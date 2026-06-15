#!/usr/bin/env python3
"""
AWS Cost Analysis & Optimization Report Generator
==================================================
Multi-Region Support - Scans ALL AWS regions where resources exist
Generates detailed Excel reports for:
- Last 6 months cost breakdown by resource (all regions)
- Idle resource identification & recommendations (per region)
- Cost savings analysis (in INR)
- Current vs projected month-end costs
- Resource usage attribution by region

Requirements: pip install boto3 pandas openpyxl xlsxwriter
"""

import boto3
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import json
import os
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

# ============================================================
# CONFIGURATION
# ============================================================

CURRENCY = "INR"
USD_TO_INR = 83.5  # Update this with current exchange rate

# Idle resource thresholds (customize as needed)
IDLE_THRESHOLDS = {
    'EC2': {
        'cpu_utilization': 5.0,
        'network_io': 1000,
        'days_without_activity': 7,
    },
    'RDS': {
        'cpu_utilization': 5.0,
        'connections': 1,
        'days_without_activity': 7,
    },
    'EBS': {
        'volume_status': 'available',
        'days_unattached': 30,
    },
    'ELB': {
        'request_count': 10,
        'days_without_requests': 7,
    },
    'EIP': {
        'association': 'unassociated',
    },
    'NAT_Gateway': {
        'bytes_processed': 1000,
        'days_low_traffic': 7,
    }
}

# ============================================================
# AWS COST EXPLORER CLIENT - MULTI REGION
# ============================================================

class AWSCostAnalyzer:
    def __init__(self, profile_name=None, access_key=None, secret_key=None, session_token=None, region='us-east-1'):
        """Initialize AWS clients and discover all regions"""
        
        # Create session with credentials if provided
        if access_key and secret_key:
            self.session = boto3.Session(
                aws_access_key_id=access_key,
                aws_secret_access_key=secret_key,
                aws_session_token=session_token
            )
        elif profile_name:
            self.session = boto3.Session(profile_name=profile_name)
        else:
            self.session = boto3.Session()
        
        # Cost Explorer is global - single client
        self.ce_client = self.session.client('ce', region_name='us-east-1')
        
        self.usd_to_inr = USD_TO_INR
        self.exchange_rate_date = datetime.now().strftime("%Y-%m-%d")
        
        # Discover all active regions
        print("🔍 Discovering AWS regions...")
        self.all_regions = self._get_all_regions()
        print(f"   Found {len(self.all_regions)} regions: {', '.join(self.all_regions)}")
        
    def _get_all_regions(self):
        """Get list of all AWS regions"""
        try:
            ec2 = self.session.client('ec2', region_name='us-east-1')
            response = ec2.describe_regions(AllRegions=False)
            return sorted([r['RegionName'] for r in response['Regions']])
        except Exception as e:
            print(f"Warning: Could not discover regions, using defaults: {e}")
            return ['us-east-1', 'us-west-2', 'eu-west-1', 'ap-south-1', 'ap-southeast-1']
    
    def _get_regional_client(self, service, region):
        """Get a regional client for a specific service"""
        return self.session.client(service, region_name=region)
    
    def convert_to_inr(self, usd_amount):
        """Convert USD to INR"""
        if usd_amount is None:
            return 0.0
        return round(float(usd_amount) * self.usd_to_inr, 2)
    
    # ============================================================
    # COST DATA COLLECTION (Global - Cost Explorer)
    # ============================================================
    
    def get_monthly_costs(self, months_back=6):
        """Get monthly cost breakdown for last N months (all regions)"""
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
                {'Type': 'DIMENSION', 'Key': 'SERVICE'},
                {'Type': 'DIMENSION', 'Key': 'LINKED_ACCOUNT'},
                {'Type': 'DIMENSION', 'Key': 'REGION'}
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
                region = group['Keys'][2]
                usd_cost = float(group['Metrics']['UnblendedCost']['Amount'])
                usage_qty = float(group['Metrics']['UsageQuantity']['Amount'])
                
                costs_data.append({
                    'Month': month,
                    'Service': service,
                    'Account_ID': account,
                    'Region': region,
                    'Cost_USD': round(usd_cost, 4),
                    'Cost_INR': self.convert_to_inr(usd_cost),
                    'Usage_Quantity': round(usage_qty, 4),
                    'Currency': 'INR'
                })
        
        return pd.DataFrame(costs_data)
    
    def get_daily_costs_current_month(self):
        """Get daily costs for current month (all regions)"""
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
                {'Type': 'DIMENSION', 'Key': 'SERVICE'},
                {'Type': 'DIMENSION', 'Key': 'REGION'}
            ]
        )
        
        daily_data = []
        for result in response.get('ResultsByTime', []):
            date = result['TimePeriod']['Start']
            for group in result.get('Groups', []):
                service = group['Keys'][0]
                region = group['Keys'][1]
                usd_cost = float(group['Metrics']['UnblendedCost']['Amount'])
                
                daily_data.append({
                    'Date': date,
                    'Service': service,
                    'Region': region,
                    'Daily_Cost_USD': round(usd_cost, 4),
                    'Daily_Cost_INR': self.convert_to_inr(usd_cost)
                })
        
        return pd.DataFrame(daily_data)
    
    def get_cost_by_usage_type(self, months_back=6):
        """Get costs broken down by usage type and region"""
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
                {'Type': 'DIMENSION', 'Key': 'SERVICE'},
                {'Type': 'DIMENSION', 'Key': 'REGION'}
            ]
        )
        
        usage_data = []
        for result in response.get('ResultsByTime', []):
            month = result['TimePeriod']['Start']
            for group in result.get('Groups', []):
                usage_type = group['Keys'][0]
                service = group['Keys'][1]
                region = group['Keys'][2]
                usd_cost = float(group['Metrics']['UnblendedCost']['Amount'])
                usage_qty = float(group['Metrics']['UsageQuantity']['Amount'])
                
                usage_data.append({
                    'Month': month,
                    'Usage_Type': usage_type,
                    'Service': service,
                    'Region': region,
                    'Cost_USD': round(usd_cost, 4),
                    'Cost_INR': self.convert_to_inr(usd_cost),
                    'Usage_Quantity': round(usage_qty, 4)
                })
        
        return pd.DataFrame(usage_data)
    
    # ============================================================
    # MULTI-REGION IDLE RESOURCE DETECTION
    # ============================================================
    
    def _scan_region_ec2(self, region):
        """Scan a single region for idle EC2 instances"""
        idle_instances = []
        try:
            ec2 = self._get_regional_client('ec2', region)
            cloudwatch = self._get_regional_client('cloudwatch', region)
            
            response = ec2.describe_instances(
                Filters=[{'Name': 'instance-state-name', 'Values': ['running']}]
            )
            
            for reservation in response.get('Reservations', []):
                for instance in reservation['Instances']:
                    instance_id = instance['InstanceId']
                    instance_type = instance['InstanceType']
                    
                    instance_name = 'Untagged'
                    for tag in instance.get('Tags', []):
                        if tag['Key'] == 'Name':
                            instance_name = tag['Value']
                            break
                    
                    # Get CPU utilization for last 7 days
                    end_time = datetime.utcnow()
                    start_time = end_time - timedelta(days=7)
                    
                    cw_response = cloudwatch.get_metric_statistics(
                        Namespace='AWS/EC2',
                        MetricName='CPUUtilization',
                        Dimensions=[{'Name': 'InstanceId', 'Value': instance_id}],
                        StartTime=start_time,
                        EndTime=end_time,
                        Period=86400,
                        Statistics=['Average']
                    )
                    
                    datapoints = cw_response.get('Datapoints', [])
                    
                    if datapoints:
                        avg_cpu = sum(dp['Average'] for dp in datapoints) / len(datapoints)
                        max_cpu = max(dp['Average'] for dp in datapoints)
                        
                        if avg_cpu < IDLE_THRESHOLDS['EC2']['cpu_utilization']:
                            monthly_cost = self._estimate_ec2_monthly_cost(instance_type, region)
                            
                            idle_instances.append({
                                'Resource_Type': 'EC2',
                                'Resource_ID': instance_id,
                                'Resource_Name': instance_name,
                                'Instance_Type': instance_type,
                                'Region': region,
                                'Status': 'Running',
                                'Avg_CPU_7d': round(avg_cpu, 2),
                                'Max_CPU_7d': round(max_cpu, 2),
                                'Monthly_Cost_USD': monthly_cost,
                                'Monthly_Cost_INR': self.convert_to_inr(monthly_cost),
                                'Recommendation': 'STOP/TERMINATE - Very low CPU utilization',
                                'Potential_Savings_INR_Monthly': self.convert_to_inr(monthly_cost),
                                'Risk_Level': 'Low' if avg_cpu < 1 else 'Medium'
                            })
        except Exception as e:
            print(f"   ⚠️  {region}: EC2 scan error - {str(e)[:60]}")
        
        return idle_instances
    
    def _scan_region_rds(self, region):
        """Scan a single region for idle RDS instances"""
        idle_rds = []
        try:
            rds = self._get_regional_client('rds', region)
            cloudwatch = self._get_regional_client('cloudwatch', region)
            
            response = rds.describe_db_instances()
            
            for db in response.get('DBInstances', []):
                db_id = db['DBInstanceIdentifier']
                engine = db['Engine']
                instance_class = db['DBInstanceClass']
                status = db['DBInstanceStatus']
                
                if status != 'available':
                    continue
                
                end_time = datetime.utcnow()
                start_time = end_time - timedelta(days=7)
                
                # CPU
                cpu_resp = cloudwatch.get_metric_statistics(
                    Namespace='AWS/RDS',
                    MetricName='CPUUtilization',
                    Dimensions=[{'Name': 'DBInstanceIdentifier', 'Value': db_id}],
                    StartTime=start_time,
                    EndTime=end_time,
                    Period=86400,
                    Statistics=['Average']
                )
                
                # Connections
                conn_resp = cloudwatch.get_metric_statistics(
                    Namespace='AWS/RDS',
                    MetricName='DatabaseConnections',
                    Dimensions=[{'Name': 'DBInstanceIdentifier', 'Value': db_id}],
                    StartTime=start_time,
                    EndTime=end_time,
                    Period=86400,
                    Statistics=['Average']
                )
                
                cpu_dp = cpu_resp.get('Datapoints', [])
                conn_dp = conn_resp.get('Datapoints', [])
                
                if cpu_dp:
                    avg_cpu = sum(dp['Average'] for dp in cpu_dp) / len(cpu_dp)
                    avg_conn = sum(dp['Average'] for dp in conn_dp) / len(conn_dp) if conn_dp else 0
                    
                    if avg_cpu < IDLE_THRESHOLDS['RDS']['cpu_utilization'] and avg_conn < IDLE_THRESHOLDS['RDS']['connections']:
                        monthly_cost = self._estimate_rds_monthly_cost(instance_class, engine)
                        
                        idle_rds.append({
                            'Resource_Type': 'RDS',
                            'Resource_ID': db_id,
                            'Resource_Name': db_id,
                            'Instance_Class': instance_class,
                            'Engine': engine,
                            'Region': region,
                            'Status': status,
                            'Avg_CPU_7d': round(avg_cpu, 2),
                            'Avg_Connections_7d': round(avg_conn, 2),
                            'Monthly_Cost_USD': monthly_cost,
                            'Monthly_Cost_INR': self.convert_to_inr(monthly_cost),
                            'Recommendation': 'STOP/DELETE - Low CPU and connections',
                            'Potential_Savings_INR_Monthly': self.convert_to_inr(monthly_cost),
                            'Risk_Level': 'Low' if avg_conn < 0.5 else 'Medium'
                        })
        except Exception as e:
            print(f"   ⚠️  {region}: RDS scan error - {str(e)[:60]}")
        
        return idle_rds
    
    def _scan_region_ebs(self, region):
        """Scan a single region for unattached EBS volumes"""
        unattached = []
        try:
            ec2 = self._get_regional_client('ec2', region)
            
            response = ec2.describe_volumes(
                Filters=[{'Name': 'status', 'Values': ['available']}]
            )
            
            for volume in response.get('Volumes', []):
                volume_id = volume['VolumeId']
                volume_type = volume['VolumeType']
                size_gb = volume['Size']
                
                monthly_cost = self._estimate_ebs_monthly_cost(volume_type, size_gb)
                
                volume_name = 'Untagged'
                for tag in volume.get('Tags', []):
                    if tag['Key'] == 'Name':
                        volume_name = tag['Value']
                        break
                
                unattached.append({
                    'Resource_Type': 'EBS Volume',
                    'Resource_ID': volume_id,
                    'Resource_Name': volume_name,
                    'Volume_Type': volume_type,
                    'Size_GB': size_gb,
                    'Region': region,
                    'Status': 'Available (Unattached)',
                    'Monthly_Cost_USD': monthly_cost,
                    'Monthly_Cost_INR': self.convert_to_inr(monthly_cost),
                    'Recommendation': 'DELETE - Not attached to any instance',
                    'Potential_Savings_INR_Monthly': self.convert_to_inr(monthly_cost),
                    'Risk_Level': 'Low',
                    'Days_Unattached': 'Unknown'
                })
        except Exception as e:
            print(f"   ⚠️  {region}: EBS scan error - {str(e)[:60]}")
        
        return unattached
    
    def _scan_region_elb(self, region):
        """Scan a single region for idle load balancers"""
        idle_lbs = []
        try:
            elb = self._get_regional_client('elbv2', region)
            cloudwatch = self._get_regional_client('cloudwatch', region)
            
            response = elb.describe_load_balancers()
            
            for lb in response.get('LoadBalancers', []):
                lb_arn = lb['LoadBalancerArn']
                lb_name = lb['LoadBalancerName']
                lb_type = lb['Type']
                
                end_time = datetime.utcnow()
                start_time = end_time - timedelta(days=7)
                
                metric_name = 'RequestCount' if lb_type == 'application' else 'ActiveFlowCount'
                namespace = 'AWS/ApplicationELB' if lb_type == 'application' else 'AWS/NetworkELB'
                
                cw_resp = cloudwatch.get_metric_statistics(
                    Namespace=namespace,
                    MetricName=metric_name,
                    Dimensions=[{'Name': 'LoadBalancer', 'Value': lb_arn.split('/')[-1]}],
                    StartTime=start_time,
                    EndTime=end_time,
                    Period=86400,
                    Statistics=['Sum']
                )
                
                datapoints = cw_resp.get('Datapoints', [])
                
                if datapoints:
                    total_requests = sum(dp['Sum'] for dp in datapoints)
                    avg_daily = total_requests / len(datapoints)
                    
                    if avg_daily < IDLE_THRESHOLDS['ELB']['request_count']:
                        monthly_cost = self._estimate_elb_monthly_cost(lb_type)
                        
                        idle_lbs.append({
                            'Resource_Type': f'{lb_type.upper()} Load Balancer',
                            'Resource_ID': lb_arn,
                            'Resource_Name': lb_name,
                            'LB_Type': lb_type,
                            'Region': region,
                            'Status': 'Active',
                            'Avg_Daily_Requests': round(avg_daily, 2),
                            'Monthly_Cost_USD': monthly_cost,
                            'Monthly_Cost_INR': self.convert_to_inr(monthly_cost),
                            'Recommendation': 'DELETE - Very low traffic',
                            'Potential_Savings_INR_Monthly': self.convert_to_inr(monthly_cost),
                            'Risk_Level': 'Low' if avg_daily == 0 else 'Medium'
                        })
        except Exception as e:
            print(f"   ⚠️  {region}: ELB scan error - {str(e)[:60]}")
        
        return idle_lbs
    
    def _scan_region_eip(self, region):
        """Scan a single region for unattached Elastic IPs"""
        unattached = []
        try:
            ec2 = self._get_regional_client('ec2', region)
            
            response = ec2.describe_addresses()
            
            for address in response.get('Addresses', []):
                if 'AssociationId' not in address:
                    allocation_id = address.get('AllocationId', 'Unknown')
                    public_ip = address['PublicIp']
                    
                    unattached.append({
                        'Resource_Type': 'Elastic IP',
                        'Resource_ID': allocation_id,
                        'Resource_Name': public_ip,
                        'Public_IP': public_ip,
                        'Region': region,
                        'Status': 'Unassociated',
                        'Monthly_Cost_USD': 3.6,
                        'Monthly_Cost_INR': self.convert_to_inr(3.6),
                        'Recommendation': 'RELEASE - Not associated with any resource',
                        'Potential_Savings_INR_Monthly': self.convert_to_inr(3.6),
                        'Risk_Level': 'Low'
                    })
        except Exception as e:
            print(f"   ⚠️  {region}: EIP scan error - {str(e)[:60]}")
        
        return unattached
    
    def scan_all_regions_idle_resources(self, max_workers=5):
        """Scan all regions for idle resources using thread pool"""
        print(f"\n🌐 Scanning {len(self.all_regions)} regions for idle resources...")
        print(f"   Using {max_workers} parallel workers\n")
        
        all_idle = {
            'EC2': [], 'RDS': [], 'EBS': [], 'ELB': [], 'EIP': []
        }
        
        def scan_region(region):
            """Scan a single region for all resource types"""
            region_results = {
                'EC2': [], 'RDS': [], 'EBS': [], 'ELB': [], 'EIP': [],
                'region': region
            }
            
            # Check if region is accessible
            try:
                ec2 = self._get_regional_client('ec2', region)
                ec2.describe_regions(RegionNames=[region])
            except Exception:
                print(f"   ❌ {region}: Region not accessible, skipping...")
                return region_results
            
            print(f"   🔎 Scanning {region}...")
            
            region_results['EC2'] = self._scan_region_ec2(region)
            region_results['RDS'] = self._scan_region_rds(region)
            region_results['EBS'] = self._scan_region_ebs(region)
            region_results['ELB'] = self._scan_region_elb(region)
            region_results['EIP'] = self._scan_region_eip(region)
            
            total_found = sum(len(v) for k, v in region_results.items() if k != 'region')
            if total_found > 0:
                print(f"   ✅ {region}: Found {total_found} idle resources")
            else:
                print(f"   ✓ {region}: No idle resources found")
            
            return region_results
        
        # Use ThreadPoolExecutor for parallel scanning
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_region = {
                executor.submit(scan_region, region): region 
                for region in self.all_regions
            }
            
            for future in as_completed(future_to_region):
                result = future.result()
                for resource_type in ['EC2', 'RDS', 'EBS', 'ELB', 'EIP']:
                    all_idle[resource_type].extend(result[resource_type])
        
        # Combine all into DataFrames
        idle_ec2 = pd.DataFrame(all_idle['EC2'])
        idle_rds = pd.DataFrame(all_idle['RDS'])
        idle_ebs = pd.DataFrame(all_idle['EBS'])
        idle_elb = pd.DataFrame(all_idle['ELB'])
        idle_eip = pd.DataFrame(all_idle['EIP'])
        
        all_idle_df = pd.concat([idle_ec2, idle_rds, idle_ebs, idle_elb, idle_eip], ignore_index=True)
        
        total_count = len(all_idle_df)
        print(f"\n📊 Total idle resources found across all regions: {total_count}")
        
        return idle_ec2, idle_rds, idle_ebs, idle_elb, idle_eip, all_idle_df
    
    # ============================================================
    # COST ESTIMATION HELPERS
    # ============================================================
    
    def _estimate_ec2_monthly_cost(self, instance_type, region):
        """Estimate monthly EC2 cost (region-aware rough approximation)"""
        base_pricing = {
            't2.micro': 8.5, 't2.small': 17, 't2.medium': 34, 't2.large': 68,
            't3.micro': 7.6, 't3.small': 15.2, 't3.medium': 30.4, 't3.large': 60.8,
            't3.xlarge': 121.6, 't3.2xlarge': 243.2,
            't3a.micro': 6.8, 't3a.small': 13.6, 't3a.medium': 27.2, 't3a.large': 54.4,
            'm5.large': 70, 'm5.xlarge': 140, 'm5.2xlarge': 280, 'm5.4xlarge': 560,
            'm5.8xlarge': 1120, 'm5.12xlarge': 1680, 'm5.16xlarge': 2240, 'm5.24xlarge': 3360,
            'm5a.large': 62, 'm5a.xlarge': 124, 'm5a.2xlarge': 248, 'm5a.4xlarge': 496,
            'm6g.large': 62, 'm6g.xlarge': 124, 'm6g.2xlarge': 248,
            'c5.large': 62, 'c5.xlarge': 124, 'c5.2xlarge': 248, 'c5.4xlarge': 496,
            'c5.9xlarge': 1116, 'c5.12xlarge': 1488, 'c5.18xlarge': 2232, 'c5.24xlarge': 2976,
            'c5a.large': 55, 'c5a.xlarge': 110, 'c5a.2xlarge': 220,
            'c6g.large': 55, 'c6g.xlarge': 110, 'c6g.2xlarge': 220,
            'r5.large': 90, 'r5.xlarge': 180, 'r5.2xlarge': 360, 'r5.4xlarge': 720,
            'r5.8xlarge': 1440, 'r5.12xlarge': 2160, 'r5.16xlarge': 2880, 'r5.24xlarge': 4320,
            'r5a.large': 80, 'r5a.xlarge': 160, 'r5a.2xlarge': 320,
            'r6g.large': 80, 'r6g.xlarge': 160, 'r6g.2xlarge': 320,
            'x1.16xlarge': 4000, 'x1.32xlarge': 8000,
            'p3.2xlarge': 3000, 'p3.8xlarge': 12000, 'p3.16xlarge': 24000,
            'g4dn.xlarge': 400, 'g4dn.2xlarge': 800, 'g4dn.4xlarge': 1600,
        }
        
        # Region multipliers (approximate)
        region_multipliers = {
            'ap-south-1': 0.90,      # Mumbai - slightly cheaper
            'ap-southeast-1': 1.00,   # Singapore
            'ap-southeast-2': 1.05,   # Sydney
            'eu-west-1': 1.00,        # Ireland
            'eu-west-2': 1.05,        # London
            'eu-central-1': 1.00,     # Frankfurt
            'us-east-1': 1.00,        # N. Virginia - baseline
            'us-east-2': 1.00,        # Ohio
            'us-west-1': 1.05,        # N. California
            'us-west-2': 1.00,        # Oregon
            'sa-east-1': 1.50,        # São Paulo - expensive
            'ca-central-1': 1.00,     # Canada
        }
        
        base = base_pricing.get(instance_type, 50)
        multiplier = region_multipliers.get(region, 1.0)
        return round(base * multiplier, 2)
    
    def _estimate_rds_monthly_cost(self, instance_class, engine):
        """Estimate monthly RDS cost"""
        pricing = {
            'db.t2.micro': 12, 'db.t2.small': 24, 'db.t2.medium': 48, 'db.t2.large': 96,
            'db.t3.micro': 11, 'db.t3.small': 22, 'db.t3.medium': 44, 'db.t3.large': 88,
            'db.t3.xlarge': 176, 'db.t3.2xlarge': 352,
            'db.t4g.micro': 9, 'db.t4g.small': 18, 'db.t4g.medium': 36, 'db.t4g.large': 72,
            'db.m5.large': 140, 'db.m5.xlarge': 280, 'db.m5.2xlarge': 560,
            'db.m5.4xlarge': 1120, 'db.m5.8xlarge': 2240, 'db.m5.12xlarge': 3360,
            'db.m5.16xlarge': 4480, 'db.m5.24xlarge': 6720,
            'db.m6g.large': 125, 'db.m6g.xlarge': 250, 'db.m6g.2xlarge': 500,
            'db.r5.large': 180, 'db.r5.xlarge': 360, 'db.r5.2xlarge': 720,
            'db.r5.4xlarge': 1440, 'db.r5.8xlarge': 2880, 'db.r5.12xlarge': 4320,
            'db.r5.16xlarge': 5760, 'db.r5.24xlarge': 8640,
            'db.r6g.large': 160, 'db.r6g.xlarge': 320, 'db.r6g.2xlarge': 640,
            'db.x2g.large': 240, 'db.x2g.xlarge': 480,
        }
        base_cost = pricing.get(instance_class, 100)
        if 'MultiAZ' in str(instance_class) or 'multi-az' in str(engine).lower():
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
        return pricing_per_gb.get(volume_type, 0.10) * size_gb * 30
    
    def _estimate_elb_monthly_cost(self, lb_type):
        """Estimate monthly ELB cost"""
        if lb_type == 'application':
            return 16.43
        elif lb_type == 'network':
            return 16.43
        else:
            return 22.50
    
    # ============================================================
    # PROJECTION & ANALYSIS
    # ============================================================
    
    def calculate_month_end_projection(self, daily_costs_df):
        """Project current month costs to month-end"""
        if daily_costs_df.empty:
            return {}
        
        today = datetime.now()
        # Handle month-end correctly
        if today.month == 12:
            next_month = today.replace(year=today.year + 1, month=1, day=1)
        else:
            next_month = today.replace(month=today.month + 1, day=1)
        days_in_month = (next_month - timedelta(days=1)).day
        
        current_day = today.day
        days_remaining = days_in_month - current_day
        
        total_so_far = daily_costs_df['Daily_Cost_INR'].sum()
        avg_daily = total_so_far / current_day if current_day > 0 else 0
        projected_total = total_so_far + (avg_daily * days_remaining)
        
        # Region + Service wise projection
        region_service_projection = {}
        for (region, service) in daily_costs_df[['Region', 'Service']].drop_duplicates().values:
            service_data = daily_costs_df[(daily_costs_df['Region'] == region) & (daily_costs_df['Service'] == service)]
            service_total = service_data['Daily_Cost_INR'].sum()
            service_avg = service_total / current_day if current_day > 0 else 0
            service_projected = service_total + (service_avg * days_remaining)
            
            key = f"{region} | {service}"
            region_service_projection[key] = {
                'Region': region,
                'Service': service,
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
            'Region_Service_Breakdown': region_service_projection
        }
    
    def calculate_savings_summary(self, idle_resources_df):
        """Calculate total potential savings from idle resources"""
        if idle_resources_df.empty:
            return {
                'Total_Idle_Resources': 0,
                'Total_Monthly_Savings_INR': 0,
                'Total_Annual_Savings_INR': 0,
                'By_Resource_Type': {},
                'By_Region': {}
            }
        
        total_monthly = idle_resources_df['Potential_Savings_INR_Monthly'].sum()
        
        by_type = idle_resources_df.groupby('Resource_Type').agg({
            'Potential_Savings_INR_Monthly': 'sum',
            'Resource_ID': 'count'
        }).reset_index().to_dict('records')
        
        by_region = idle_resources_df.groupby('Region').agg({
            'Potential_Savings_INR_Monthly': 'sum',
            'Resource_ID': 'count'
        }).reset_index().to_dict('records')
        
        return {
            'Total_Idle_Resources': len(idle_resources_df),
            'Total_Monthly_Savings_INR': round(total_monthly, 2),
            'Total_Annual_Savings_INR': round(total_monthly * 12, 2),
            'By_Resource_Type': by_type,
            'By_Region': by_region
        }
    
    # ============================================================
    # EXCEL REPORT GENERATION
    # ============================================================
    
    def generate_excel_report(self, output_file='aws_cost_report.xlsx'):
        """Generate comprehensive Excel report with multi-region data"""
        print("=" * 70)
        print("AWS COST ANALYSIS & OPTIMIZATION REPORT - MULTI-REGION")
        print("=" * 70)
        print(f"Exchange Rate: 1 USD = {self.usd_to_inr} INR (as of {self.exchange_rate_date})")
        print(f"Regions Scanned: {len(self.all_regions)}")
        print(f"Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("-" * 70)
        
        # Collect all data
        print("\n[1/5] Collecting monthly cost data (all regions)...")
        monthly_costs = self.get_monthly_costs(6)
        
        print("[2/5] Collecting daily costs for projection (all regions)...")
        daily_costs = self.get_daily_costs_current_month()
        
        print("[3/5] Collecting usage type breakdown (all regions)...")
        usage_types = self.get_cost_by_usage_type(6)
        
        print("[4/5] Detecting idle resources across ALL regions...")
        idle_ec2, idle_rds, idle_ebs, idle_elb, idle_eip, all_idle = self.scan_all_regions_idle_resources()
        
        print("[5/5] Calculating projections and savings...")
        projection = self.calculate_month_end_projection(daily_costs)
        savings = self.calculate_savings_summary(all_idle)
        
        print("\n[6/5] Generating Excel report...")
        
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Define formats
            header_format = workbook.add_format({
                'bold': True, 'bg_color': '#366092', 'font_color': 'white',
                'border': 1, 'align': 'center', 'valign': 'vcenter'
            })
            header_format_left = workbook.add_format({
                'bold': True, 'bg_color': '#366092', 'font_color': 'white',
                'border': 1, 'align': 'left', 'valign': 'vcenter'
            })
            money_format = workbook.add_format({'num_format': '₹#,##0.00', 'border': 1})
            money_format_red = workbook.add_format({
                'num_format': '₹#,##0.00', 'border': 1, 'font_color': '#C00000', 'bold': True
            })
            money_format_green = workbook.add_format({
                'num_format': '₹#,##0.00', 'border': 1, 'font_color': '#00B050', 'bold': True
            })
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
            title_format = workbook.add_format({
                'bold': True, 'font_size': 14, 'bg_color': '#366092',
                'font_color': 'white', 'border': 1, 'align': 'center'
            })
            
            # ============================================================
            # SHEET 1: EXECUTIVE SUMMARY
            # ============================================================
            summary_data = [
                ['AWS COST ANALYSIS - MULTI REGION REPORT', ''],
                ['', ''],
                ['Report Details', ''],
                ['Total Regions Scanned', len(self.all_regions)],
                ['Regions', ', '.join(self.all_regions)],
                ['Exchange Rate (1 USD)', f'{self.usd_to_inr} INR'],
                ['Rate Date', self.exchange_rate_date],
                ['Report Generated', datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                ['', ''],
            ]
            
            if projection:
                summary_data.extend([
                    ['CURRENT MONTH PROJECTION (All Regions)', ''],
                    ['Current Day of Month', projection['Current_Month_Day']],
                    ['Days in Month', projection['Days_In_Month']],
                    ['Days Remaining', projection['Days_Remaining']],
                    ['Current Cost (INR)', projection['Current_Cost_INR']],
                    ['Projected Month-End Cost (INR)', projection['Projected_Month_End_INR']],
                    ['Remaining Days Cost (INR)', projection['Remaining_Days_Cost_INR']],
                    ['Average Daily Cost (INR)', projection['Average_Daily_Cost_INR']],
                    ['', ''],
                ])
            
            summary_data.extend([
                ['IDLE RESOURCE SAVINGS POTENTIAL (All Regions)', ''],
                ['Total Idle Resources Found', savings['Total_Idle_Resources']],
                ['Total Monthly Savings Potential (INR)', savings['Total_Monthly_Savings_INR']],
                ['Total Annual Savings Potential (INR)', savings['Total_Annual_Savings_INR']],
                ['', ''],
            ])
            
            # Add region-wise savings breakdown
            if savings.get('By_Region'):
                summary_data.append(['SAVINGS BY REGION', ''])
                for region_data in savings['By_Region']:
                    summary_data.append([
                        f"  {region_data['Region']}",
                        f"{region_data['Resource_ID']} resources | ₹{region_data['Potential_Savings_INR_Monthly']:,.2f}/month"
                    ])
                summary_data.append(['', ''])
            
            summary_df = pd.DataFrame(summary_data, columns=['Metric', 'Value'])
            summary_df.to_excel(writer, sheet_name='Executive Summary', index=False)
            
            summary_sheet = writer.sheets['Executive Summary']
            summary_sheet.set_column('A:A', 45)
            summary_sheet.set_column('B:B', 45)
            
            for row_num in range(len(summary_data)):
                if summary_data[row_num][0] in [
                    'AWS COST ANALYSIS - MULTI REGION REPORT',
                    'Report Details',
                    'CURRENT MONTH PROJECTION (All Regions)',
                    'IDLE RESOURCE SAVINGS POTENTIAL (All Regions)',
                    'SAVINGS BY REGION'
                ]:
                    summary_sheet.write(row_num, 0, summary_data[row_num][0], title_format if 'AWS' in summary_data[row_num][0] else header_format)
                    summary_sheet.write(row_num, 1, '', title_format if 'AWS' in summary_data[row_num][0] else header_format)
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
            # SHEET 2: MONTHLY COST TREND (By Region)
            # ============================================================
            if not monthly_costs.empty:
                pivot_monthly = monthly_costs.pivot_table(
                    index=['Service', 'Region', 'Account_ID'],
                    columns='Month',
                    values='Cost_INR',
                    aggfunc='sum',
                    fill_value=0
                ).reset_index()
                
                pivot_monthly.to_excel(writer, sheet_name='Monthly Cost Trend', index=False)
                trend_sheet = writer.sheets['Monthly Cost Trend']
                trend_sheet.set_column('A:A', 35)
                trend_sheet.set_column('B:B', 18)
                trend_sheet.set_column('C:C', 18)
                
                for col_num in range(3, len(pivot_monthly.columns)):
                    trend_sheet.set_column(col_num, col_num, 18, money_format)
                
                for col_num, col_name in enumerate(pivot_monthly.columns):
                    trend_sheet.write(0, col_num, col_name, header_format)
            
            # ============================================================
            # SHEET 3: DAILY COST TRACKING (By Region)
            # ============================================================
            if not daily_costs.empty:
                daily_costs.to_excel(writer, sheet_name='Daily Cost Tracking', index=False)
                daily_sheet = writer.sheets['Daily Cost Tracking']
                daily_sheet.set_column('A:A', 15, date_format)
                daily_sheet.set_column('B:B', 35)
                daily_sheet.set_column('C:C', 18)
                daily_sheet.set_column('D:D', 18, money_format)
                daily_sheet.set_column('E:E', 18, money_format)
                
                for col_num in range(len(daily_costs.columns)):
                    daily_sheet.write(0, col_num, daily_costs.columns[col_num], header_format)
            
            # ============================================================
            # SHEET 4: COST BY USAGE TYPE (By Region)
            # ============================================================
            if not usage_types.empty:
                usage_types.to_excel(writer, sheet_name='Cost by Usage Type', index=False)
                usage_sheet = writer.sheets['Cost by Usage Type']
                usage_sheet.set_column('A:A', 15)
                usage_sheet.set_column('B:B', 40)
                usage_sheet.set_column('C:C', 30)
                usage_sheet.set_column('D:D', 18)
                usage_sheet.set_column('E:E', 18, money_format)
                usage_sheet.set_column('F:F', 18, money_format)
                usage_sheet.set_column('G:G', 18, cell_format)
                
                for col_num in range(len(usage_types.columns)):
                    usage_sheet.write(0, col_num, usage_types.columns[col_num], header_format)
            
            # ============================================================
            # SHEET 5: IDLE EC2 INSTANCES (All Regions)
            # ============================================================
            if not idle_ec2.empty:
                idle_ec2.to_excel(writer, sheet_name='Idle EC2 Instances', index=False)
                ec2_sheet = writer.sheets['Idle EC2 Instances']
                ec2_sheet.set_column('A:A', 15)
                ec2_sheet.set_column('B:B', 25)
                ec2_sheet.set_column('C:C', 25)
                ec2_sheet.set_column('D:D', 18)
                ec2_sheet.set_column('E:E', 18)
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
            # SHEET 6: IDLE RDS INSTANCES (All Regions)
            # ============================================================
            if not idle_rds.empty:
                idle_rds.to_excel(writer, sheet_name='Idle RDS Instances', index=False)
                rds_sheet = writer.sheets['Idle RDS Instances']
                for col_num in range(len(idle_rds.columns)):
                    rds_sheet.write(0, col_num, idle_rds.columns[col_num], header_format)
                rds_sheet.set_column('L:L', 18, money_format_green)
            
            # ============================================================
            # SHEET 7: UNATTACHED EBS VOLUMES (All Regions)
            # ============================================================
            if not idle_ebs.empty:
                idle_ebs.to_excel(writer, sheet_name='Unattached EBS Volumes', index=False)
                ebs_sheet = writer.sheets['Unattached EBS Volumes']
                for col_num in range(len(idle_ebs.columns)):
                    ebs_sheet.write(0, col_num, idle_ebs.columns[col_num], header_format)
                ebs_sheet.set_column('J:J', 18, money_format_green)
            
            # ============================================================
            # SHEET 8: IDLE LOAD BALANCERS (All Regions)
            # ============================================================
            if not idle_elb.empty:
                idle_elb.to_excel(writer, sheet_name='Idle Load Balancers', index=False)
                elb_sheet = writer.sheets['Idle Load Balancers']
                for col_num in range(len(idle_elb.columns)):
                    elb_sheet.write(0, col_num, idle_elb.columns[col_num], header_format)
                elb_sheet.set_column('J:J', 18, money_format_green)
            
            # ============================================================
            # SHEET 9: UNATTACHED ELASTIC IPs (All Regions)
            # ============================================================
            if not idle_eip.empty:
                idle_eip.to_excel(writer, sheet_name='Unattached Elastic IPs', index=False)
                eip_sheet = writer.sheets['Unattached Elastic IPs']
                for col_num in range(len(idle_eip.columns)):
                    eip_sheet.write(0, col_num, idle_eip.columns[col_num], header_format)
                eip_sheet.set_column('H:H', 18, money_format_green)
            
            # ============================================================
            # SHEET 10: ALL IDLE RESOURCES COMBINED (All Regions)
            # ============================================================
            if not all_idle.empty:
                all_idle.to_excel(writer, sheet_name='All Idle Resources', index=False)
                all_sheet = writer.sheets['All Idle Resources']
                all_sheet.set_column('A:A', 20)
                all_sheet.set_column('B:B', 30)
                all_sheet.set_column('C:C', 25)
                all_sheet.set_column('D:D', 18)  # Region column
                
                savings_col = None
                for idx, col in enumerate(all_idle.columns):
                    if 'Savings' in col:
                        savings_col = idx
                        break
                
                if savings_col is not None:
                    all_sheet.set_column(savings_col, savings_col, 22, money_format_green)
                
                for col_num in range(len(all_idle.columns)):
                    all_sheet.write(0, col_num, all_idle.columns[col_num], header_format)
                
                # Conditional formatting for risk levels
                risk_col = None
                for idx, col in enumerate(all_idle.columns):
                    if col == 'Risk_Level':
                        risk_col = idx
                        break
                
                if risk_col is not None:
                    all_sheet.conditional_format(1, risk_col, len(all_idle), risk_col, {
                        'type': 'cell', 'criteria': 'equal to', 'value': '"High"', 'format': warning_format
                    })
                    all_sheet.conditional_format(1, risk_col, len(all_idle), risk_col, {
                        'type': 'cell', 'criteria': 'equal to', 'value': '"Low"', 'format': success_format
                    })
            
            # ============================================================
            # SHEET 11: MONTH-END PROJECTION (By Region)
            # ============================================================
            if projection and projection.get('Region_Service_Breakdown'):
                projection_data = []
                for key, data in projection['Region_Service_Breakdown'].items():
                    projection_data.append([
                        data['Region'],
                        data['Service'],
                        data['Current_Cost_INR'],
                        data['Remaining_Days_Cost_INR'],
                        data['Projected_Total_INR']
                    ])
                
                projection_df = pd.DataFrame(projection_data, columns=[
                    'Region', 'Service', 'Current Cost (INR)', 'Remaining Days (INR)', 'Projected Month-End (INR)'
                ])
                projection_df.to_excel(writer, sheet_name='Month-End Projection', index=False)
                
                proj_sheet = writer.sheets['Month-End Projection']
                proj_sheet.set_column('A:A', 18)
                proj_sheet.set_column('B:B', 35)
                proj_sheet.set_column('C:C', 20, money_format)
                proj_sheet.set_column('D:D', 20, money_format)
                proj_sheet.set_column('E:E', 22, money_format_red)
                
                for col_num in range(len(projection_df.columns)):
                    proj_sheet.write(0, col_num, projection_df.columns[col_num], header_format)
            
            # ============================================================
            # SHEET 12: SAVINGS SUMMARY (By Type & Region)
            # ============================================================
            savings_data = []
            savings_data.append(['Resource Type', 'Region', 'Count', 'Monthly Savings (INR)', 'Annual Savings (INR)', 'Action Required'])
            
            if not all_idle.empty:
                grouped = all_idle.groupby(['Resource_Type', 'Region']).agg({
                    'Resource_ID': 'count',
                    'Potential_Savings_INR_Monthly': 'sum'
                }).reset_index()
                
                for _, row in grouped.iterrows():
                    savings_data.append([
                        row['Resource_Type'],
                        row['Region'],
                        row['Resource_ID'],
                        round(row['Potential_Savings_INR_Monthly'], 2),
                        round(row['Potential_Savings_INR_Monthly'] * 12, 2),
                        'Review and Delete/Stop'
                    ])
            
            savings_data.append(['', '', '', '', '', ''])
            savings_data.append(['TOTAL', 'All Regions', savings['Total_Idle_Resources'], 
                               savings['Total_Monthly_Savings_INR'], 
                               savings['Total_Annual_Savings_INR'], ''])
            
            savings_df = pd.DataFrame(savings_data[1:], columns=savings_data[0])
            savings_df.to_excel(writer, sheet_name='Savings Summary', index=False)
            
            savings_sheet = writer.sheets['Savings Summary']
            savings_sheet.set_column('A:A', 22)
            savings_sheet.set_column('B:B', 18)
            savings_sheet.set_column('C:C', 12, center_format)
            savings_sheet.set_column('D:D', 22, money_format_green)
            savings_sheet.set_column('E:E', 22, money_format_green)
            savings_sheet.set_column('F:F', 25)
            
            for col_num in range(len(savings_df.columns)):
                savings_sheet.write(0, col_num, savings_df.columns[col_num], header_format)
            
            # Highlight total row
            total_row = len(savings_df)
            total_format = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
            for col_num in range(len(savings_df.columns)):
                savings_sheet.write(total_row, col_num, savings_df.iloc[-1, col_num], total_format)
            
            # ============================================================
            # SHEET 13: REGION SUMMARY
            # ============================================================
            if not monthly_costs.empty:
                region_summary = monthly_costs.groupby('Region').agg({
                    'Cost_INR': 'sum',
                    'Service': 'nunique'
                }).reset_index()
                region_summary.columns = ['Region', 'Total Cost 6mo (INR)', 'Services Used']
                region_summary = region_summary.sort_values('Total Cost 6mo (INR)', ascending=False)
                
                region_summary.to_excel(writer, sheet_name='Region Summary', index=False)
                region_sheet = writer.sheets['Region Summary']
                region_sheet.set_column('A:A', 18)
                region_sheet.set_column('B:B', 22, money_format)
                region_sheet.set_column('C:C', 15, center_format)
                
                for col_num in range(len(region_summary.columns)):
                    region_sheet.write(0, col_num, region_summary.columns[col_num], header_format)
        
        print(f"\n✅ Report generated successfully: {os.path.abspath(output_file)}")
        print(f"\n📊 Report Contents (13 Sheets):")
        print(f"   1. Executive Summary - Overview, projections, savings by region")
        print(f"   2. Monthly Cost Trend - 6-month cost history by service & region")
        print(f"   3. Daily Cost Tracking - Current month daily breakdown by region")
        print(f"   4. Cost by Usage Type - Detailed usage-based costing by region")
        print(f"   5. Idle EC2 Instances - Underutilized compute (all regions)")
        print(f"   6. Idle RDS Instances - Underutilized databases (all regions)")
        print(f"   7. Unattached EBS Volumes - Unused storage (all regions)")
        print(f"   8. Idle Load Balancers - Low-traffic LBs (all regions)")
        print(f"   9. Unattached Elastic IPs - Unassociated IPs (all regions)")
        print(f"  10. All Idle Resources - Combined list with region & risk levels")
        print(f"  11. Month-End Projection - Current vs projected by region & service")
        print(f"  12. Savings Summary - Total savings by resource type & region")
        print(f"  13. Region Summary - Cost ranking of all regions")
        
        return output_file


# ============================================================
# MAIN EXECUTION
# ============================================================

def main():
    """Main function to run the Multi-Region AWS Cost Analyzer"""
    import argparse
    
    parser = argparse.ArgumentParser(description='AWS Multi-Region Cost Analysis & Optimization Report')
    parser.add_argument('--profile', type=str, help='AWS profile name (from ~/.aws/credentials)')
    parser.add_argument('--access-key', type=str, help='AWS Access Key ID')
    parser.add_argument('--secret-key', type=str, help='AWS Secret Access Key')
    parser.add_argument('--session-token', type=str, help='AWS Session Token (for temporary credentials)')
    parser.add_argument('--output', type=str, default='aws_cost_report.xlsx', help='Output Excel file name')
    parser.add_argument('--exchange-rate', type=float, default=83.5, help='USD to INR exchange rate')
    parser.add_argument('--months', type=int, default=6, help='Number of months to analyze')
    parser.add_argument('--workers', type=int, default=5, help='Parallel workers for region scanning')
    
    args = parser.parse_args()
    
    global USD_TO_INR
    USD_TO_INR = args.exchange_rate
    
    print("\n" + "=" * 70)
    print("AWS MULTI-REGION COST ANALYSIS & OPTIMIZATION TOOL")
    print("=" * 70)
    
    try:
        analyzer = AWSCostAnalyzer(
            profile_name=args.profile,
            access_key=args.access_key,
            secret_key=args.secret_key,
            session_token=args.session_token
        )
        analyzer.generate_excel_report(output_file=args.output)
    except Exception as e:
        print(f"\n❌ Error: {e}")
        print("\nTroubleshooting:")
        print("1. Ensure AWS credentials are valid")
        print("2. Ensure Cost Explorer is enabled in AWS Console")
        print("3. Check IAM permissions for all services in all regions")
        print("4. Install required packages: pip install -r requirements.txt")
        raise


if __name__ == '__main__':
    main()
