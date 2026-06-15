#!/usr/bin/env python3
"""
AWS Cost Analysis & Optimization Report Generator
==================================================
Generates detailed Excel reports for:
- Last 6 months cost breakdown by resource
- Idle resource identification & recommendations
- Cost savings analysis (in INR)
- Current vs projected month-end costs
- Resource usage attribution

Requirements: pip install boto3 pandas openpyxl xlsxwriter
"""

import boto3
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import json
import os
from collections import defaultdict

# ============================================================
# CONFIGURATION
# ============================================================

# AWS Cost Explorer API has limits - adjust as needed
DAYS_PER_REQUEST = 365  # Max for GetCostAndUsage
CURRENCY = "INR"
USD_TO_INR = 83.5  # Update this with current exchange rate

# Idle resource thresholds (customize as needed)
IDLE_THRESHOLDS = {
    'EC2': {
        'cpu_utilization': 5.0,      # % - below this is considered idle
        'network_io': 1000,          # bytes - below this is idle
        'days_without_activity': 7,  # days
    },
    'RDS': {
        'cpu_utilization': 5.0,
        'connections': 1,
        'days_without_activity': 7,
    },
    'EBS': {
        'volume_status': 'available',  # Not attached
        'days_unattached': 30,
    },
    'ELB': {
        'request_count': 10,         # per day
        'days_without_requests': 7,
    },
    'EIP': {
        'association': 'unassociated',
    },
    'NAT_Gateway': {
        'bytes_processed': 1000,     # Very low traffic
        'days_low_traffic': 7,
    }
}

# ============================================================
# AWS COST EXPLORER CLIENT
# ============================================================

class AWSCostAnalyzer:
    def __init__(self, profile_name=None, region='us-east-1'):
        """Initialize AWS clients"""
        session = boto3.Session(profile_name=profile_name) if profile_name else boto3.Session()
        
        self.ce_client = session.client('ce', region_name=region)  # Cost Explorer
        self.ec2_client = session.client('ec2', region_name=region)
        self.rds_client = session.client('rds', region_name=region)
        self.cloudwatch_client = session.client('cloudwatch', region_name=region)
        self.elb_client = session.client('elbv2', region_name=region)
        
        self.usd_to_inr = USD_TO_INR
        self.exchange_rate_date = datetime.now().strftime("%Y-%m-%d")
        
    def convert_to_inr(self, usd_amount):
        """Convert USD to INR"""
        if usd_amount is None:
            return 0.0
        return round(float(usd_amount) * self.usd_to_inr, 2)
    
    # ============================================================
    # COST DATA COLLECTION
    # ============================================================
    
    def get_monthly_costs(self, months_back=6):
        """Get monthly cost breakdown for last N months"""
        end_date = datetime.now().replace(day=1)  # First day of current month
        start_date = end_date - relativedelta(months=months_back)
        
        # Adjust to get full months
        start_date = start_date.replace(day=1)
        
        response = self.ce_client.get_cost_and_usage(
            TimePeriod={
                'Start': start_date.strftime('%Y-%m-%d'),
                'End': end_date.strftime('%Y-%m-%d')
            },
            Granularity='MONTHLY',
            Metrics=['UnblendedCost', 'UsageQuantity'],
            GroupBy=[
                {'Type': 'DIMENSION', 'Key': 'SERVICE'},
                {'Type': 'DIMENSION', 'Key': 'LINKED_ACCOUNT'}
            ],
            Filter={
                'Not': {
                    'Dimensions': {
                        'Key': 'RECORD_TYPE',
                        'Values': ['Credit', 'Refund', 'UpfrontReservationFee']
                    }
                }
            }
        )
        
        costs_data = []
        for result in response.get('ResultsByTime', []):
            month = result['TimePeriod']['Start']
            for group in result.get('Groups', []):
                service = group['Keys'][0]
                account = group['Keys'][1]
                usd_cost = float(group['Metrics']['UnblendedCost']['Amount'])
                usage_qty = float(group['Metrics']['UsageQuantity']['Amount'])
                
                costs_data.append({
                    'Month': month,
                    'Service': service,
                    'Account_ID': account,
                    'Cost_USD': round(usd_cost, 4),
                    'Cost_INR': self.convert_to_inr(usd_cost),
                    'Usage_Quantity': round(usage_qty, 4),
                    'Currency': 'INR'
                })
        
        return pd.DataFrame(costs_data)
    
    def get_daily_costs_current_month(self):
        """Get daily costs for current month to project month-end"""
        today = datetime.now()
        start_date = today.replace(day=1)
        end_date = today + timedelta(days=1)
        
        response = self.ce_client.get_cost_and_usage(
            TimePeriod={
                'Start': start_date.strftime('%Y-%m-%d'),
                'End': end_date.strftime('%Y-%m-%d')
            },
            Granularity='DAILY',
            Metrics=['UnblendedCost'],
            GroupBy=[
                {'Type': 'DIMENSION', 'Key': 'SERVICE'}
            ]
        )
        
        daily_data = []
        for result in response.get('ResultsByTime', []):
            date = result['TimePeriod']['Start']
            daily_cost = 0
            for group in result.get('Groups', []):
                service = group['Keys'][0]
                usd_cost = float(group['Metrics']['UnblendedCost']['Amount'])
                daily_cost += usd_cost
                
                daily_data.append({
                    'Date': date,
                    'Service': service,
                    'Daily_Cost_USD': round(usd_cost, 4),
                    'Daily_Cost_INR': self.convert_to_inr(usd_cost)
                })
        
        return pd.DataFrame(daily_data)
    
    def get_resource_level_costs(self, months_back=6):
        """Get cost allocation by resource tags (if tagging enabled)"""
        end_date = datetime.now().replace(day=1)
        start_date = end_date - relativedelta(months=months_back)
        start_date = start_date.replace(day=1)
        
        # Try to get costs by resource tags
        try:
            response = self.ce_client.get_cost_and_usage(
                TimePeriod={
                    'Start': start_date.strftime('%Y-%m-%d'),
                    'End': end_date.strftime('%Y-%m-%d')
                },
                Granularity='MONTHLY',
                Metrics=['UnblendedCost'],
                GroupBy=[
                    {'Type': 'TAG', 'Key': 'Name'},
                    {'Type': 'DIMENSION', 'Key': 'SERVICE'}
                ]
            )
            
            resource_data = []
            for result in response.get('ResultsByTime', []):
                month = result['TimePeriod']['Start']
                for group in result.get('Groups', []):
                    resource_name = group['Keys'][0] if group['Keys'][0] else 'Untagged'
                    service = group['Keys'][1]
                    usd_cost = float(group['Metrics']['UnblendedCost']['Amount'])
                    
                    resource_data.append({
                        'Month': month,
                        'Resource_Name': resource_name,
                        'Service': service,
                        'Cost_USD': round(usd_cost, 4),
                        'Cost_INR': self.convert_to_inr(usd_cost)
                    })
            
            return pd.DataFrame(resource_data)
        except Exception as e:
            print(f"Warning: Could not get resource-level costs: {e}")
            return pd.DataFrame()
    
    def get_cost_by_usage_type(self, months_back=6):
        """Get costs broken down by usage type (helps identify idle resources)"""
        end_date = datetime.now().replace(day=1)
        start_date = end_date - relativedelta(months=months_back)
        start_date = start_date.replace(day=1)
        
        response = self.ce_client.get_cost_and_usage(
            TimePeriod={
                'Start': start_date.strftime('%Y-%m-%d'),
                'End': end_date.strftime('%Y-%m-%d')
            },
            Granularity='MONTHLY',
            Metrics=['UnblendedCost', 'UsageQuantity'],
            GroupBy=[
                {'Type': 'DIMENSION', 'Key': 'USAGE_TYPE'},
                {'Type': 'DIMENSION', 'Key': 'SERVICE'}
            ]
        )
        
        usage_data = []
        for result in response.get('ResultsByTime', []):
            month = result['TimePeriod']['Start']
            for group in result.get('Groups', []):
                usage_type = group['Keys'][0]
                service = group['Keys'][1]
                usd_cost = float(group['Metrics']['UnblendedCost']['Amount'])
                usage_qty = float(group['Metrics']['UsageQuantity']['Amount'])
                
                usage_data.append({
                    'Month': month,
                    'Usage_Type': usage_type,
                    'Service': service,
                    'Cost_USD': round(usd_cost, 4),
                    'Cost_INR': self.convert_to_inr(usd_cost),
                    'Usage_Quantity': round(usage_qty, 4)
                })
        
        return pd.DataFrame(usage_data)
    
    # ============================================================
    # IDLE RESOURCE DETECTION
    # ============================================================
    
    def get_idle_ec2_instances(self):
        """Find EC2 instances with low CPU utilization"""
        idle_instances = []
        
        try:
            # Get all running instances
            ec2_resource = boto3.resource('ec2')
            instances = ec2_resource.instances.filter(
                Filters=[{'Name': 'instance-state-name', 'Values': ['running']}]
            )
            
            for instance in instances:
                instance_id = instance.id
                instance_name = 'Unknown'
                for tag in instance.tags or []:
                    if tag['Key'] == 'Name':
                        instance_name = tag['Value']
                        break
                
                # Get CPU utilization for last 7 days
                end_time = datetime.utcnow()
                start_time = end_time - timedelta(days=7)
                
                response = self.cloudwatch_client.get_metric_statistics(
                    Namespace='AWS/EC2',
                    MetricName='CPUUtilization',
                    Dimensions=[{'Name': 'InstanceId', 'Value': instance_id}],
                    StartTime=start_time,
                    EndTime=end_time,
                    Period=86400,  # Daily
                    Statistics=['Average']
                )
                
                datapoints = sorted(response.get('Datapoints', []), key=lambda x: x['Timestamp'])
                
                if datapoints:
                    avg_cpu = sum(dp['Average'] for dp in datapoints) / len(datapoints)
                    max_cpu = max(dp['Average'] for dp in datapoints)
                    
                    if avg_cpu < IDLE_THRESHOLDS['EC2']['cpu_utilization']:
                        # Get instance cost estimate
                        monthly_cost = self._estimate_ec2_monthly_cost(instance.instance_type, instance.placement['AvailabilityZone'])
                        
                        idle_instances.append({
                            'Resource_Type': 'EC2',
                            'Resource_ID': instance_id,
                            'Resource_Name': instance_name,
                            'Instance_Type': instance.instance_type,
                            'Region': instance.placement['AvailabilityZone'],
                            'Status': 'Running',
                            'Avg_CPU_7d': round(avg_cpu, 2),
                            'Max_CPU_7d': round(max_cpu, 2),
                            'Monthly_Cost_USD': monthly_cost,
                            'Monthly_Cost_INR': self.convert_to_inr(monthly_cost),
                            'Recommendation': 'STOP/TERMINATE - Very low CPU utilization',
                            'Potential_Savings_INR_Monthly': self.convert_to_inr(monthly_cost),
                            'Risk_Level': 'Low' if avg_cpu < 1 else 'Medium'
                        })
                else:
                    # No CloudWatch data - might be newly launched or monitoring disabled
                    pass
                    
        except Exception as e:
            print(f"Error checking EC2 instances: {e}")
        
        return pd.DataFrame(idle_instances)
    
    def get_idle_rds_instances(self):
        """Find RDS instances with low activity"""
        idle_rds = []
        
        try:
            response = self.rds_client.describe_db_instances()
            
            for db in response['DBInstances']:
                db_id = db['DBInstanceIdentifier']
                engine = db['Engine']
                instance_class = db['DBInstanceClass']
                
                # Get CloudWatch metrics
                end_time = datetime.utcnow()
                start_time = end_time - timedelta(days=7)
                
                # CPU Utilization
                cpu_response = self.cloudwatch_client.get_metric_statistics(
                    Namespace='AWS/RDS',
                    MetricName='CPUUtilization',
                    Dimensions=[{'Name': 'DBInstanceIdentifier', 'Value': db_id}],
                    StartTime=start_time,
                    EndTime=end_time,
                    Period=86400,
                    Statistics=['Average']
                )
                
                # Database Connections
                conn_response = self.cloudwatch_client.get_metric_statistics(
                    Namespace='AWS/RDS',
                    MetricName='DatabaseConnections',
                    Dimensions=[{'Name': 'DBInstanceIdentifier', 'Value': db_id}],
                    StartTime=start_time,
                    EndTime=end_time,
                    Period=86400,
                    Statistics=['Average']
                )
                
                cpu_datapoints = cpu_response.get('Datapoints', [])
                conn_datapoints = conn_response.get('Datapoints', [])
                
                if cpu_datapoints:
                    avg_cpu = sum(dp['Average'] for dp in cpu_datapoints) / len(cpu_datapoints)
                    avg_conn = sum(dp['Average'] for dp in conn_datapoints) / len(conn_datapoints) if conn_datapoints else 0
                    
                    if avg_cpu < IDLE_THRESHOLDS['RDS']['cpu_utilization'] and avg_conn < IDLE_THRESHOLDS['RDS']['connections']:
                        monthly_cost = self._estimate_rds_monthly_cost(instance_class, engine)
                        
                        idle_rds.append({
                            'Resource_Type': 'RDS',
                            'Resource_ID': db_id,
                            'Resource_Name': db_id,
                            'Instance_Class': instance_class,
                            'Engine': engine,
                            'Status': db['DBInstanceStatus'],
                            'Avg_CPU_7d': round(avg_cpu, 2),
                            'Avg_Connections_7d': round(avg_conn, 2),
                            'Monthly_Cost_USD': monthly_cost,
                            'Monthly_Cost_INR': self.convert_to_inr(monthly_cost),
                            'Recommendation': 'STOP/DELETE - Low CPU and connections',
                            'Potential_Savings_INR_Monthly': self.convert_to_inr(monthly_cost),
                            'Risk_Level': 'Low' if avg_conn < 0.5 else 'Medium'
                        })
                        
        except Exception as e:
            print(f"Error checking RDS instances: {e}")
        
        return pd.DataFrame(idle_rds)
    
    def get_unattached_ebs_volumes(self):
        """Find EBS volumes that are not attached to any instance"""
        unattached_volumes = []
        
        try:
            response = self.ec2_client.describe_volumes(
                Filters=[{'Name': 'status', 'Values': ['available']}]
            )
            
            for volume in response['Volumes']:
                volume_id = volume['VolumeId']
                volume_type = volume['VolumeType']
                size_gb = volume['Size']
                
                # Calculate monthly cost
                monthly_cost = self._estimate_ebs_monthly_cost(volume_type, size_gb)
                
                # Get volume name from tags
                volume_name = 'Untagged'
                for tag in volume.get('Tags', []):
                    if tag['Key'] == 'Name':
                        volume_name = tag['Value']
                        break
                
                unattached_volumes.append({
                    'Resource_Type': 'EBS Volume',
                    'Resource_ID': volume_id,
                    'Resource_Name': volume_name,
                    'Volume_Type': volume_type,
                    'Size_GB': size_gb,
                    'Status': 'Available (Unattached)',
                    'Monthly_Cost_USD': monthly_cost,
                    'Monthly_Cost_INR': self.convert_to_inr(monthly_cost),
                    'Recommendation': 'DELETE - Not attached to any instance',
                    'Potential_Savings_INR_Monthly': self.convert_to_inr(monthly_cost),
                    'Risk_Level': 'Low',
                    'Days_Unattached': 'Unknown'  # Would need to check CloudTrail for exact date
                })
                
        except Exception as e:
            print(f"Error checking EBS volumes: {e}")
        
        return pd.DataFrame(unattached_volumes)
    
    def get_idle_load_balancers(self):
        """Find load balancers with very low request count"""
        idle_lbs = []
        
        try:
            # Application and Network Load Balancers
            response = self.elb_client.describe_load_balancers()
            
            for lb in response['LoadBalancers']:
                lb_arn = lb['LoadBalancerArn']
                lb_name = lb['LoadBalancerName']
                lb_type = lb['Type']
                
                end_time = datetime.utcnow()
                start_time = end_time - timedelta(days=7)
                
                # Get request count
                metric_name = 'RequestCount' if lb_type == 'application' else 'ActiveFlowCount'
                
                response_cw = self.cloudwatch_client.get_metric_statistics(
                    Namespace='AWS/ApplicationELB' if lb_type == 'application' else 'AWS/NetworkELB',
                    MetricName=metric_name,
                    Dimensions=[{'Name': 'LoadBalancer', 'Value': lb_arn.split('/')[-1]}],
                    StartTime=start_time,
                    EndTime=end_time,
                    Period=86400,
                    Statistics=['Sum']
                )
                
                datapoints = response_cw.get('Datapoints', [])
                
                if datapoints:
                    total_requests = sum(dp['Sum'] for dp in datapoints)
                    avg_daily = total_requests / len(datapoints) if datapoints else 0
                    
                    if avg_daily < IDLE_THRESHOLDS['ELB']['request_count']:
                        monthly_cost = self._estimate_elb_monthly_cost(lb_type)
                        
                        idle_lbs.append({
                            'Resource_Type': f'{lb_type.upper()} Load Balancer',
                            'Resource_ID': lb_arn,
                            'Resource_Name': lb_name,
                            'LB_Type': lb_type,
                            'Status': 'Active',
                            'Avg_Daily_Requests': round(avg_daily, 2),
                            'Monthly_Cost_USD': monthly_cost,
                            'Monthly_Cost_INR': self.convert_to_inr(monthly_cost),
                            'Recommendation': 'DELETE - Very low traffic',
                            'Potential_Savings_INR_Monthly': self.convert_to_inr(monthly_cost),
                            'Risk_Level': 'Low' if avg_daily == 0 else 'Medium'
                        })
                        
        except Exception as e:
            print(f"Error checking Load Balancers: {e}")
        
        return pd.DataFrame(idle_lbs)
    
    def get_unattached_eips(self):
        """Find Elastic IPs that are not associated with any instance"""
        unattached_eips = []
        
        try:
            response = self.ec2_client.describe_addresses()
            
            for address in response['Addresses']:
                if 'AssociationId' not in address:
                    allocation_id = address.get('AllocationId', 'Unknown')
                    public_ip = address['PublicIp']
                    
                    unattached_eips.append({
                        'Resource_Type': 'Elastic IP',
                        'Resource_ID': allocation_id,
                        'Resource_Name': public_ip,
                        'Public_IP': public_ip,
                        'Status': 'Unassociated',
                        'Monthly_Cost_USD': 3.6,  # Standard AWS charge for unattached EIP
                        'Monthly_Cost_INR': self.convert_to_inr(3.6),
                        'Recommendation': 'RELEASE - Not associated with any resource',
                        'Potential_Savings_INR_Monthly': self.convert_to_inr(3.6),
                        'Risk_Level': 'Low'
                    })
                    
        except Exception as e:
            print(f"Error checking Elastic IPs: {e}")
        
        return pd.DataFrame(unattached_eips)
    
    # ============================================================
    # COST ESTIMATION HELPERS
    # ============================================================
    
    def _estimate_ec2_monthly_cost(self, instance_type, region):
        """Estimate monthly EC2 cost (rough approximation)"""
        # These are rough estimates - actual costs vary by region and usage
        pricing = {
            't2.micro': 8.5, 't2.small': 17, 't2.medium': 34, 't2.large': 68,
            't3.micro': 7.6, 't3.small': 15.2, 't3.medium': 30.4, 't3.large': 60.8,
            't3.xlarge': 121.6, 't3.2xlarge': 243.2,
            'm5.large': 70, 'm5.xlarge': 140, 'm5.2xlarge': 280,
            'm5.4xlarge': 560, 'm5.8xlarge': 1120, 'm5.12xlarge': 1680,
            'c5.large': 62, 'c5.xlarge': 124, 'c5.2xlarge': 248,
            'c5.4xlarge': 496, 'c5.9xlarge': 1116,
            'r5.large': 90, 'r5.xlarge': 180, 'r5.2xlarge': 360,
            'r5.4xlarge': 720, 'r5.8xlarge': 1440,
        }
        return pricing.get(instance_type, 50)  # Default $50 if unknown
    
    def _estimate_rds_monthly_cost(self, instance_class, engine):
        """Estimate monthly RDS cost"""
        pricing = {
            'db.t2.micro': 12, 'db.t2.small': 24, 'db.t2.medium': 48,
            'db.t3.micro': 11, 'db.t3.small': 22, 'db.t3.medium': 44,
            'db.t3.large': 88, 'db.t3.xlarge': 176,
            'db.m5.large': 140, 'db.m5.xlarge': 280,
            'db.m5.2xlarge': 560, 'db.m5.4xlarge': 1120,
            'db.r5.large': 180, 'db.r5.xlarge': 360,
            'db.r5.2xlarge': 720, 'db.r5.4xlarge': 1440,
        }
        base_cost = pricing.get(instance_class, 100)
        # Multi-AZ doubles the cost
        if 'MultiAZ' in str(instance_class):
            base_cost *= 2
        return base_cost
    
    def _estimate_ebs_monthly_cost(self, volume_type, size_gb):
        """Estimate monthly EBS cost"""
        pricing_per_gb = {
            'gp2': 0.10, 'gp3': 0.08,
            'io1': 0.125, 'io2': 0.125,
            'st1': 0.045, 'sc1': 0.025,
            'standard': 0.05
        }
        return pricing_per_gb.get(volume_type, 0.10) * size_gb * 30  # 30 days
    
    def _estimate_elb_monthly_cost(self, lb_type):
        """Estimate monthly ELB cost"""
        if lb_type == 'application':
            return 16.43  # ALB base cost + LCUs
        elif lb_type == 'network':
            return 16.43  # NLB base cost
        else:
            return 22.50  # Classic LB
    
    # ============================================================
    # PROJECTION & ANALYSIS
    # ============================================================
    
    def calculate_month_end_projection(self, daily_costs_df):
        """Project current month costs to month-end"""
        if daily_costs_df.empty:
            return {}
        
        today = datetime.now()
        days_in_month = (today.replace(month=today.month+1, day=1) - timedelta(days=1)).day
        current_day = today.day
        days_remaining = days_in_month - current_day
        
        # Calculate average daily cost so far
        total_so_far = daily_costs_df['Daily_Cost_INR'].sum()
        avg_daily = total_so_far / current_day if current_day > 0 else 0
        
        projected_total = total_so_far + (avg_daily * days_remaining)
        
        # Service-wise projection
        service_projection = {}
        for service in daily_costs_df['Service'].unique():
            service_data = daily_costs_df[daily_costs_df['Service'] == service]
            service_total = service_data['Daily_Cost_INR'].sum()
            service_avg = service_total / current_day if current_day > 0 else 0
            service_projected = service_total + (service_avg * days_remaining)
            service_projection[service] = {
                'Current_Cost_INR': round(service_total, 2),
                'Projected_Total_INR': round(service_projected, 2),
                'Remaining_Days_Cost_INR': round(service_avg * days_remaining, 2)
            }
        
        return {
            'Current_Month_Day': current_day,
            'Days_In_Month': days_in_month,
            'Days_Remaining': days_remaining,
            'Current_Cost_INR': round(total_so_far, 2),
            'Projected_Month_End_INR': round(projected_total, 2),
            'Remaining_Days_Cost_INR': round(avg_daily * days_remaining, 2),
            'Average_Daily_Cost_INR': round(avg_daily, 2),
            'Service_Breakdown': service_projection
        }
    
    def calculate_savings_summary(self, idle_resources_df):
        """Calculate total potential savings from idle resources"""
        if idle_resources_df.empty:
            return {
                'Total_Idle_Resources': 0,
                'Total_Monthly_Savings_INR': 0,
                'Total_Annual_Savings_INR': 0,
                'By_Resource_Type': {}
            }
        
        total_monthly = idle_resources_df['Potential_Savings_INR_Monthly'].sum()
        
        by_type = idle_resources_df.groupby('Resource_Type').agg({
            'Potential_Savings_INR_Monthly': 'sum',
            'Resource_ID': 'count'
        }).to_dict()
        
        return {
            'Total_Idle_Resources': len(idle_resources_df),
            'Total_Monthly_Savings_INR': round(total_monthly, 2),
            'Total_Annual_Savings_INR': round(total_monthly * 12, 2),
            'By_Resource_Type': by_type
        }
    
    # ============================================================
    # EXCEL REPORT GENERATION
    # ============================================================
    
    def generate_excel_report(self, output_file='aws_cost_report.xlsx'):
        """Generate comprehensive Excel report"""
        print("=" * 60)
        print("AWS COST ANALYSIS & OPTIMIZATION REPORT")
        print("=" * 60)
        print(f"Exchange Rate: 1 USD = {self.usd_to_inr} INR (as of {self.exchange_rate_date})")
        print(f"Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("-" * 60)
        
        # Collect all data
        print("\n[1/7] Collecting monthly cost data...")
        monthly_costs = self.get_monthly_costs(6)
        
        print("[2/7] Collecting daily costs for projection...")
        daily_costs = self.get_daily_costs_current_month()
        
        print("[3/7] Collecting resource-level costs...")
        resource_costs = self.get_resource_level_costs(6)
        
        print("[4/7] Collecting usage type breakdown...")
        usage_types = self.get_cost_by_usage_type(6)
        
        print("[5/7] Detecting idle resources...")
        idle_ec2 = self.get_idle_ec2_instances()
        idle_rds = self.get_idle_rds_instances()
        idle_ebs = self.get_unattached_ebs_volumes()
        idle_elb = self.get_idle_load_balancers()
        idle_eip = self.get_unattached_eips()
        
        # Combine all idle resources
        all_idle = pd.concat([idle_ec2, idle_rds, idle_ebs, idle_elb, idle_eip], ignore_index=True)
        
        print("[6/7] Calculating projections and savings...")
        projection = self.calculate_month_end_projection(daily_costs)
        savings = self.calculate_savings_summary(all_idle)
        
        print("[7/7] Generating Excel report...")
        
        # Create Excel writer
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Define formats
            header_format = workbook.add_format({
                'bold': True, 'bg_color': '#366092', 'font_color': 'white',
                'border': 1, 'align': 'center', 'valign': 'vcenter'
            })
            money_format = workbook.add_format({'num_format': '₹#,##0.00', 'border': 1})
            money_format_red = workbook.add_format({
                'num_format': '₹#,##0.00', 'border': 1, 'font_color': '#C00000', 'bold': True
            })
            money_format_green = workbook.add_format({
                'num_format': '₹#,##0.00', 'border': 1, 'font_color': '#00B050', 'bold': True
            })
            percent_format = workbook.add_format({'num_format': '0.00%', 'border': 1})
            date_format = workbook.add_format({'num_format': 'YYYY-MM-DD', 'border': 1})
            cell_format = workbook.add_format({'border': 1, 'align': 'left'})
            center_format = workbook.add_format({'border': 1, 'align': 'center'})
            warning_format = workbook.add_format({
                'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'bold': True
            })
            success_format = workbook.add_format({
                'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1, 'bold': True
            })
            info_format = workbook.add_format({
                'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'border': 1, 'bold': True
            })
            
            # ============================================================
            # SHEET 1: EXECUTIVE SUMMARY
            # ============================================================
            summary_data = []
            
            # Current month info
            if projection:
                summary_data.extend([
                    ['', ''],
                    ['CURRENT MONTH PROJECTION', ''],
                    ['Current Day of Month', projection['Current_Month_Day']],
                    ['Days in Month', projection['Days_In_Month']],
                    ['Days Remaining', projection['Days_Remaining']],
                    ['Current Cost (INR)', projection['Current_Cost_INR']],
                    ['Projected Month-End Cost (INR)', projection['Projected_Month_End_INR']],
                    ['Remaining Days Cost (INR)', projection['Remaining_Days_Cost_INR']],
                    ['Average Daily Cost (INR)', projection['Average_Daily_Cost_INR']],
                    ['', ''],
                ])
            
            # Savings summary
            summary_data.extend([
                ['IDLE RESOURCE SAVINGS POTENTIAL', ''],
                ['Total Idle Resources Found', savings['Total_Idle_Resources']],
                ['Total Monthly Savings Potential (INR)', savings['Total_Monthly_Savings_INR']],
                ['Total Annual Savings Potential (INR)', savings['Total_Annual_Savings_INR']],
                ['', ''],
                ['EXCHANGE RATE', ''],
                ['1 USD = INR', self.usd_to_inr],
                ['Rate Date', self.exchange_rate_date],
            ])
            
            summary_df = pd.DataFrame(summary_data, columns=['Metric', 'Value'])
            summary_df.to_excel(writer, sheet_name='Executive Summary', index=False)
            
            # Format summary sheet
            summary_sheet = writer.sheets['Executive Summary']
            summary_sheet.set_column('A:A', 40)
            summary_sheet.set_column('B:B', 25)
            
            # Apply formatting to summary
            for row_num in range(len(summary_data)):
                if summary_data[row_num][0] in ['CURRENT MONTH PROJECTION', 'IDLE RESOURCE SAVINGS POTENTIAL', 'EXCHANGE RATE']:
                    summary_sheet.write(row_num, 0, summary_data[row_num][0], header_format)
                    summary_sheet.write(row_num, 1, '', header_format)
                elif 'INR' in str(summary_data[row_num][0]):
                    summary_sheet.write(row_num, 0, summary_data[row_num][0], cell_format)
                    if 'Savings' in str(summary_data[row_num][0]) or 'Potential' in str(summary_data[row_num][0]):
                        summary_sheet.write(row_num, 1, summary_data[row_num][1], money_format_green)
                    else:
                        summary_sheet.write(row_num, 1, summary_data[row_num][1], money_format)
                else:
                    summary_sheet.write(row_num, 0, summary_data[row_num][0], cell_format)
                    summary_sheet.write(row_num, 1, summary_data[row_num][1], center_format)
            
            # ============================================================
            # SHEET 2: MONTHLY COST TREND
            # ============================================================
            if not monthly_costs.empty:
                # Pivot for better view
                pivot_monthly = monthly_costs.pivot_table(
                    index=['Service', 'Account_ID'],
                    columns='Month',
                    values='Cost_INR',
                    aggfunc='sum',
                    fill_value=0
                ).reset_index()
                
                pivot_monthly.to_excel(writer, sheet_name='Monthly Cost Trend', index=False)
                trend_sheet = writer.sheets['Monthly Cost Trend']
                trend_sheet.set_column('A:A', 35)
                trend_sheet.set_column('B:B', 18)
                
                # Auto-adjust column widths for month columns
                for col_num in range(2, len(pivot_monthly.columns)):
                    trend_sheet.set_column(col_num, col_num, 18, money_format)
                
                # Add header formatting
                for col_num, col_name in enumerate(pivot_monthly.columns):
                    trend_sheet.write(0, col_num, col_name, header_format)
            
            # ============================================================
            # SHEET 3: DAILY COST TRACKING
            # ============================================================
            if not daily_costs.empty:
                daily_costs.to_excel(writer, sheet_name='Daily Cost Tracking', index=False)
                daily_sheet = writer.sheets['Daily Cost Tracking']
                daily_sheet.set_column('A:A', 15, date_format)
                daily_sheet.set_column('B:B', 35)
                daily_sheet.set_column('C:C', 18, money_format)
                daily_sheet.set_column('D:D', 18, money_format)
                
                for col_num in range(len(daily_costs.columns)):
                    daily_sheet.write(0, col_num, daily_costs.columns[col_num], header_format)
            
            # ============================================================
            # SHEET 4: COST BY USAGE TYPE
            # ============================================================
            if not usage_types.empty:
                usage_types.to_excel(writer, sheet_name='Cost by Usage Type', index=False)
                usage_sheet = writer.sheets['Cost by Usage Type']
                usage_sheet.set_column('A:A', 15)
                usage_sheet.set_column('B:B', 40)
                usage_sheet.set_column('C:C', 30)
                usage_sheet.set_column('D:D', 18, money_format)
                usage_sheet.set_column('E:E', 18, money_format)
                usage_sheet.set_column('F:F', 18, cell_format)
                
                for col_num in range(len(usage_types.columns)):
                    usage_sheet.write(0, col_num, usage_types.columns[col_num], header_format)
            
            # ============================================================
            # SHEET 5: IDLE EC2 INSTANCES
            # ============================================================
            if not idle_ec2.empty:
                idle_ec2.to_excel(writer, sheet_name='Idle EC2 Instances', index=False)
                ec2_sheet = writer.sheets['Idle EC2 Instances']
                ec2_sheet.set_column('A:A', 15)
                ec2_sheet.set_column('B:B', 25)
                ec2_sheet.set_column('C:C', 25)
                ec2_sheet.set_column('D:D', 18)
                ec2_sheet.set_column('E:E', 20)
                ec2_sheet.set_column('F:F', 15)
                ec2_sheet.set_column('G:G', 15, cell_format)
                ec2_sheet.set_column('H:H', 15, cell_format)
                ec2_sheet.set_column('I:I', 18, money_format)
                ec2_sheet.set_column('J:J', 18, money_format)
                ec2_sheet.set_column('K:K', 40)
                ec2_sheet.set_column('L:L', 18, money_format_green)
                ec2_sheet.set_column('M:M', 12)
                
                for col_num in range(len(idle_ec2.columns)):
                    ec2_sheet.write(0, col_num, idle_ec2.columns[col_num], header_format)
            
            # ============================================================
            # SHEET 6: IDLE RDS INSTANCES
            # ============================================================
            if not idle_rds.empty:
                idle_rds.to_excel(writer, sheet_name='Idle RDS Instances', index=False)
                rds_sheet = writer.sheets['Idle RDS Instances']
                for col_num in range(len(idle_rds.columns)):
                    rds_sheet.write(0, col_num, idle_rds.columns[col_num], header_format)
                rds_sheet.set_column('L:L', 18, money_format_green)
            
            # ============================================================
            # SHEET 7: UNATTACHED EBS VOLUMES
            # ============================================================
            if not idle_ebs.empty:
                idle_ebs.to_excel(writer, sheet_name='Unattached EBS Volumes', index=False)
                ebs_sheet = writer.sheets['Unattached EBS Volumes']
                for col_num in range(len(idle_ebs.columns)):
                    ebs_sheet.write(0, col_num, idle_ebs.columns[col_num], header_format)
                ebs_sheet.set_column('J:J', 18, money_format_green)
            
            # ============================================================
            # SHEET 8: IDLE LOAD BALANCERS
            # ============================================================
            if not idle_elb.empty:
                idle_elb.to_excel(writer, sheet_name='Idle Load Balancers', index=False)
                elb_sheet = writer.sheets['Idle Load Balancers']
                for col_num in range(len(idle_elb.columns)):
                    elb_sheet.write(0, col_num, idle_elb.columns[col_num], header_format)
                elb_sheet.set_column('J:J', 18, money_format_green)
            
            # ============================================================
            # SHEET 9: UNATTACHED ELASTIC IPs
            # ============================================================
            if not idle_eip.empty:
                idle_eip.to_excel(writer, sheet_name='Unattached Elastic IPs', index=False)
                eip_sheet = writer.sheets['Unattached Elastic IPs']
                for col_num in range(len(idle_eip.columns)):
                    eip_sheet.write(0, col_num, idle_eip.columns[col_num], header_format)
                eip_sheet.set_column('H:H', 18, money_format_green)
            
            # ============================================================
            # SHEET 10: ALL IDLE RESOURCES COMBINED
            # ============================================================
            if not all_idle.empty:
                all_idle.to_excel(writer, sheet_name='All Idle Resources', index=False)
                all_sheet = writer.sheets['All Idle Resources']
                all_sheet.set_column('A:A', 20)
                all_sheet.set_column('B:B', 30)
                all_sheet.set_column('C:C', 25)
                
                # Find the savings column index
                savings_col = None
                for idx, col in enumerate(all_idle.columns):
                    if 'Savings' in col:
                        savings_col = idx
                        break
                
                if savings_col is not None:
                    all_sheet.set_column(savings_col, savings_col, 20, money_format_green)
                
                for col_num in range(len(all_idle.columns)):
                    all_sheet.write(0, col_num, all_idle.columns[col_num], header_format)
                
                # Add conditional formatting for risk levels
                risk_col = None
                for idx, col in enumerate(all_idle.columns):
                    if col == 'Risk_Level':
                        risk_col = idx
                        break
                
                if risk_col is not None:
                    all_sheet.conditional_format(1, risk_col, len(all_idle), risk_col, {
                        'type': 'cell',
                        'criteria': 'equal to',
                        'value': '"High"',
                        'format': warning_format
                    })
                    all_sheet.conditional_format(1, risk_col, len(all_idle), risk_col, {
                        'type': 'cell',
                        'criteria': 'equal to',
                        'value': '"Low"',
                        'format': success_format
                    })
            
            # ============================================================
            # SHEET 11: MONTH-END PROJECTION DETAIL
            # ============================================================
            if projection and projection.get('Service_Breakdown'):
                projection_data = []
                for service, data in projection['Service_Breakdown'].items():
                    projection_data.append([
                        service,
                        data['Current_Cost_INR'],
                        data['Remaining_Days_Cost_INR'],
                        data['Projected_Total_INR']
                    ])
                
                projection_df = pd.DataFrame(projection_data, columns=[
                    'Service', 'Current Cost (INR)', 'Remaining Days (INR)', 'Projected Month-End (INR)'
                ])
                projection_df.to_excel(writer, sheet_name='Month-End Projection', index=False)
                
                proj_sheet = writer.sheets['Month-End Projection']
                proj_sheet.set_column('A:A', 35)
                proj_sheet.set_column('B:B', 20, money_format)
                proj_sheet.set_column('C:C', 20, money_format)
                proj_sheet.set_column('D:D', 22, money_format_red)
                
                for col_num in range(len(projection_df.columns)):
                    proj_sheet.write(0, col_num, projection_df.columns[col_num], header_format)
            
            # ============================================================
            # SHEET 12: SAVINGS SUMMARY
            # ============================================================
            savings_data = []
            savings_data.append(['Resource Type', 'Count', 'Monthly Savings (INR)', 'Annual Savings (INR)', 'Action Required'])
            
            if not all_idle.empty:
                grouped = all_idle.groupby('Resource_Type').agg({
                    'Resource_ID': 'count',
                    'Potential_Savings_INR_Monthly': 'sum'
                }).reset_index()
                
                for _, row in grouped.iterrows():
                    savings_data.append([
                        row['Resource_Type'],
                        row['Resource_ID'],
                        round(row['Potential_Savings_INR_Monthly'], 2),
                        round(row['Potential_Savings_INR_Monthly'] * 12, 2),
                        'Review and Delete/Stop'
                    ])
            
            savings_data.append(['', '', '', '', ''])
            savings_data.append(['TOTAL', savings['Total_Idle_Resources'], savings['Total_Monthly_Savings_INR'], 
                               savings['Total_Annual_Savings_INR'], ''])
            
            savings_df = pd.DataFrame(savings_data[1:], columns=savings_data[0])
            savings_df.to_excel(writer, sheet_name='Savings Summary', index=False)
            
            savings_sheet = writer.sheets['Savings Summary']
            savings_sheet.set_column('A:A', 25)
            savings_sheet.set_column('B:B', 12, center_format)
            savings_sheet.set_column('C:C', 22, money_format_green)
            savings_sheet.set_column('D:D', 22, money_format_green)
            savings_sheet.set_column('E:E', 25)
            
            for col_num in range(len(savings_df.columns)):
                savings_sheet.write(0, col_num, savings_df.columns[col_num], header_format)
            
            # Highlight total row
            total_row = len(savings_df)
            for col_num in range(len(savings_df.columns)):
                savings_sheet.write(total_row, col_num, savings_df.iloc[-1, col_num], 
                                  workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1}))
        
        print(f"\n✅ Report generated successfully: {os.path.abspath(output_file)}")
        print(f"\n📊 Report Contents:")
        print(f"   • Executive Summary - Overview of costs and savings")
        print(f"   • Monthly Cost Trend - 6-month cost history by service")
        print(f"   • Daily Cost Tracking - Current month daily breakdown")
        print(f"   • Cost by Usage Type - Detailed usage-based costing")
        print(f"   • Idle EC2 Instances - Underutilized compute resources")
        print(f"   • Idle RDS Instances - Underutilized database resources")
        print(f"   • Unattached EBS Volumes - Unused storage volumes")
        print(f"   • Idle Load Balancers - Low-traffic load balancers")
        print(f"   • Unattached Elastic IPs - Unassociated IP addresses")
        print(f"   • All Idle Resources - Combined idle resource list")
        print(f"   • Month-End Projection - Current vs projected costs")
        print(f"   • Savings Summary - Total potential savings analysis")
        
        return output_file


# ============================================================
# MAIN EXECUTION
# ============================================================

def main():
    """Main function to run the AWS Cost Analyzer"""
    import argparse
    
    parser = argparse.ArgumentParser(description='AWS Cost Analysis & Optimization Report')
    parser.add_argument('--profile', type=str, help='AWS profile name (from ~/.aws/credentials)')
    parser.add_argument('--region', type=str, default='us-east-1', help='AWS region')
    parser.add_argument('--output', type=str, default='aws_cost_report.xlsx', help='Output Excel file name')
    parser.add_argument('--exchange-rate', type=float, default=83.5, help='USD to INR exchange rate')
    parser.add_argument('--months', type=int, default=6, help='Number of months to analyze')
    
    args = parser.parse_args()
    
    # Update exchange rate
    global USD_TO_INR
    USD_TO_INR = args.exchange_rate
    
    print("\n" + "=" * 60)
    print("AWS COST ANALYSIS & OPTIMIZATION TOOL")
    print("=" * 60)
    
    try:
        analyzer = AWSCostAnalyzer(profile_name=args.profile, region=args.region)
        analyzer.generate_excel_report(output_file=args.output)
    except Exception as e:
        print(f"\n❌ Error: {e}")
        print("\nTroubleshooting:")
        print("1. Ensure AWS credentials are configured (aws configure)")
        print("2. Ensure Cost Explorer is enabled in AWS Console")
        print("3. Check IAM permissions for Cost Explorer, EC2, RDS, CloudWatch, ELB")
        print("4. Install required packages: pip install boto3 pandas openpyxl xlsxwriter")
        raise


if __name__ == '__main__':
    main()
