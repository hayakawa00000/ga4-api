from flask import Flask, jsonify, request
from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import RunReportRequest
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import requests as http_requests
import os
import json
from datetime import datetime

app = Flask(__name__)

SERVICE_ACCOUNT_JSON = os.environ.get('SERVICE_ACCOUNT_JSON')
DEFAULT_PROPERTY_ID = os.environ.get('GA4_PROPERTY_ID')

GSC_REFRESH_TOKEN = os.environ.get('GSC_REFRESH_TOKEN')
GSC_CLIENT_ID = os.environ.get('GSC_CLIENT_ID')
GSC_CLIENT_SECRET = os.environ.get('GSC_CLIENT_SECRET')

GOOGLE_ADS_REFRESH_TOKEN = os.environ.get('GOOGLE_ADS_REFRESH_TOKEN')
GOOGLE_ADS_CLIENT_ID = os.environ.get('GOOGLE_ADS_CLIENT_ID')
GOOGLE_ADS_CLIENT_SECRET = os.environ.get('GOOGLE_ADS_CLIENT_SECRET')
GOOGLE_ADS_DEVELOPER_TOKEN = os.environ.get('GOOGLE_ADS_DEVELOPER_TOKEN')

def get_ads_access_token():
    response = http_requests.post('https://oauth2.googleapis.com/token', data={
        'client_id': GOOGLE_ADS_CLIENT_ID,
        'client_secret': GOOGLE_ADS_CLIENT_SECRET,
        'refresh_token': GOOGLE_ADS_REFRESH_TOKEN,
        'grant_type': 'refresh_token'
    })
    return response.json().get('access_token')

def query_google_ads(customer_id, query):
    access_token = get_ads_access_token()
    url = f"https://googleads.googleapis.com/v15/customers/{customer_id}/googleAds:search"
    headers = {
        'Authorization': f'Bearer {access_token}',
        'developer-token': GOOGLE_ADS_DEVELOPER_TOKEN,
        'login-customer-id': customer_id,
        'Content-Type': 'application/json'
    }
    response = http_requests.post(url, headers=headers, json={'query': query})
    if response.status_code != 200:
        return {'error': response.text}
    return response.json()

@app.route('/')
def home():
    return jsonify({"status": "GA4 & GSC & Google Ads API is running"})

@app.route('/health')
def health():
    return jsonify({"status": "ok"})

@app.route('/google-ads/debug')
def debug_google_ads():
    try:
        customer_id = request.args.get('customer_id')

        token_response = http_requests.post('https://oauth2.googleapis.com/token', data={
            'client_id': GOOGLE_ADS_CLIENT_ID,
            'client_secret': GOOGLE_ADS_CLIENT_SECRET,
            'refresh_token': GOOGLE_ADS_REFRESH_TOKEN,
            'grant_type': 'refresh_token'
        })
        token_data = token_response.json()
        access_token = token_data.get('access_token')

        if not access_token:
            return jsonify({"step": "token_failed", "token_response": token_data})

        url = f"https://googleads.googleapis.com/v17/customers/{customer_id}/googleAds:search"
        headers = {
            'Authorization': f'Bearer {access_token}',
            'developer-token': GOOGLE_ADS_DEVELOPER_TOKEN,
            'login-customer-id': customer_id,
            'Content-Type': 'application/json'
        }
        api_response = http_requests.post(url, headers=headers, json={
            'query': 'SELECT campaign.name FROM campaign LIMIT 1'
        })

        return jsonify({
            "step": "api_called",
            "status_code": api_response.status_code,
            "url_used": url,
            "customer_id": customer_id,
            "developer_token_set": bool(GOOGLE_ADS_DEVELOPER_TOKEN),
            "response_text": api_response.text[:1000]
        })

    except Exception as e:
        return jsonify({"error": str(e)})

@app.route('/ga4/sessions')
def get_sessions():
    try:
        property_id = request.args.get('property_id', DEFAULT_PROPERTY_ID)
        start_date = request.args.get('start_date', '7daysAgo')
        end_date = request.args.get('end_date', 'today')

        if not property_id:
            return jsonify({"success": False, "error": "GA4_PROPERTY_ID が設定されていません"}), 500
        if not SERVICE_ACCOUNT_JSON:
            return jsonify({"success": False, "error": "SERVICE_ACCOUNT_JSON が設定されていません"}), 500

        service_account_info = json.loads(SERVICE_ACCOUNT_JSON)
        credentials = service_account.Credentials.from_service_account_info(
            service_account_info,
            scopes=['https://www.googleapis.com/auth/analytics.readonly']
        )
        client = BetaAnalyticsDataClient(credentials=credentials)
        request_obj = RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[{"start_date": start_date, "end_date": end_date}],
            metrics=[{"name": "sessions"}]
        )
        response = client.run_report(request_obj)
        sessions = 0
        if response.rows:
            sessions = response.rows[0].metric_values[0].value
        return jsonify({"success": True, "sessions": int(sessions), "start_date": start_date, "end_date": end_date})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/ga4/comprehensive')
def get_comprehensive():
    try:
        property_id = request.args.get('property_id', DEFAULT_PROPERTY_ID)
        start_date = request.args.get('start_date', '7daysAgo')
        end_date = request.args.get('end_date', 'today')

        if not property_id:
            return jsonify({"success": False, "error": "GA4_PROPERTY_ID が設定されていません"}), 500
        if not SERVICE_ACCOUNT_JSON:
            return jsonify({"success": False, "error": "SERVICE_ACCOUNT_JSON が設定されていません"}), 500

        service_account_info = json.loads(SERVICE_ACCOUNT_JSON)
        credentials = service_account.Credentials.from_service_account_info(
            service_account_info,
            scopes=['https://www.googleapis.com/auth/analytics.readonly']
        )
        client = BetaAnalyticsDataClient(credentials=credentials)

        summary_response = client.run_report(RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[{"start_date": start_date, "end_date": end_date}],
            metrics=[
                {"name": "sessions"}, {"name": "activeUsers"},
                {"name": "screenPageViews"}, {"name": "engagementRate"},
                {"name": "bounceRate"}, {"name": "averageSessionDuration"},
                {"name": "screenPageViewsPerSession"}, {"name": "keyEvents"}
            ]
        ))
        summary = {}
        if summary_response.rows:
            row = summary_response.rows[0]
            summary = {
                "sessions": int(row.metric_values[0].value),
                "active_users": int(row.metric_values[1].value),
                "pageviews": int(row.metric_values[2].value),
                "engagement_rate": float(row.metric_values[3].value),
                "bounce_rate": float(row.metric_values[4].value),
                "average_session_duration": float(row.metric_values[5].value),
                "pageviews_per_session": float(row.metric_values[6].value),
                "key_events": int(row.metric_values[7].value)
            }

        source_response = client.run_report(RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[{"start_date": start_date, "end_date": end_date}],
            dimensions=[{"name": "sessionSource"}, {"name": "sessionMedium"}],
            metrics=[{"name": "sessions"}, {"name": "activeUsers"}],
            order_bys=[{"metric": {"metric_name": "sessions"}, "desc": True}],
            limit=10
        ))
        traffic_sources = []
        for row in source_response.rows:
            traffic_sources.append({
                "source": row.dimension_values[0].value,
                "medium": row.dimension_values[1].value,
                "sessions": int(row.metric_values[0].value),
                "users": int(row.metric_values[1].value)
            })

        device_response = client.run_report(RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[{"start_date": start_date, "end_date": end_date}],
            dimensions=[{"name": "deviceCategory"}],
            metrics=[{"name": "sessions"}, {"name": "activeUsers"}, {"name": "engagementRate"}],
            order_bys=[{"metric": {"metric_name": "sessions"}, "desc": True}]
        ))
        devices = []
        for row in device_response.rows:
            devices.append({
                "device": row.dimension_values[0].value,
                "sessions": int(row.metric_values[0].value),
                "users": int(row.metric_values[1].value),
                "engagement_rate": float(row.metric_values[2].value)
            })

        page_response = client.run_report(RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[{"start_date": start_date, "end_date": end_date}],
            dimensions=[{"name": "pagePath"}],
            metrics=[{"name": "screenPageViews"}, {"name": "activeUsers"}, {"name": "averageSessionDuration"}],
            order_bys=[{"metric": {"metric_name": "screenPageViews"}, "desc": True}],
            limit=20
        ))
        pages = []
        for row in page_response.rows:
            pages.append({
                "page_path": row.dimension_values[0].value,
                "pageviews": int(row.metric_values[0].value),
                "users": int(row.metric_values[1].value),
                "avg_session_duration": float(row.metric_values[2].value)
            })

        city_response = client.run_report(RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[{"start_date": start_date, "end_date": end_date}],
            dimensions=[{"name": "city"}],
            metrics=[{"name": "sessions"}, {"name": "activeUsers"}],
            order_bys=[{"metric": {"metric_name": "sessions"}, "desc": True}],
            limit=10
        ))
        cities = []
        for row in city_response.rows:
            cities.append({
                "city": row.dimension_values[0].value,
                "sessions": int(row.metric_values[0].value),
                "users": int(row.metric_values[1].value)
            })

        landing_response = client.run_report(RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[{"start_date": start_date, "end_date": end_date}],
            dimensions=[{"name": "landingPage"}],
            metrics=[{"name": "sessions"}, {"name": "bounceRate"}, {"name": "engagementRate"}],
            order_bys=[{"metric": {"metric_name": "sessions"}, "desc": True}],
            limit=10
        ))
        landing_pages = []
        for row in landing_response.rows:
            landing_pages.append({
                "landing_page": row.dimension_values[0].value,
                "sessions": int(row.metric_values[0].value),
                "bounce_rate": float(row.metric_values[1].value),
                "engagement_rate": float(row.metric_values[2].value)
            })

        event_response = client.run_report(RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[{"start_date": start_date, "end_date": end_date}],
            dimensions=[{"name": "eventName"}],
            metrics=[{"name": "eventCount"}],
            order_bys=[{"metric": {"metric_name": "eventCount"}, "desc": True}],
            limit=10
        ))
        events = []
        for row in event_response.rows:
            events.append({
                "event_name": row.dimension_values[0].value,
                "event_count": int(row.metric_values[0].value)
            })

        return jsonify({
            "success": True,
            "property_id": property_id,
            "start_date": start_date,
            "end_date": end_date,
            "summary": summary,
            "traffic_sources": traffic_sources,
            "devices": devices,
            "pages": pages,
            "cities": cities,
            "landing_pages": landing_pages,
            "events": events
        })

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/gsc/queries')
def get_gsc_queries():
    try:
        site_url = request.args.get('site_url')
        start_date = request.args.get('start_date', '2025-01-01')
        end_date = request.args.get('end_date', '2025-01-31')
        limit = int(request.args.get('limit', 20))

        if not site_url:
            return jsonify({"success": False, "error": "site_url が必要です"}), 400
        if not all([GSC_REFRESH_TOKEN, GSC_CLIENT_ID, GSC_CLIENT_SECRET]):
            return jsonify({"success": False, "error": "GSC環境変数が設定されていません"}), 500

        creds = Credentials(
            token=None, refresh_token=GSC_REFRESH_TOKEN,
            token_uri='https://oauth2.googleapis.com/token',
            client_id=GSC_CLIENT_ID, client_secret=GSC_CLIENT_SECRET,
            scopes=['https://www.googleapis.com/auth/webmasters.readonly']
        )
        service = build('searchconsole', 'v1', credentials=creds)

        summary_response = service.searchanalytics().query(
            siteUrl=site_url,
            body={'startDate': start_date, 'endDate': end_date}
        ).execute()

        summary = {'total_clicks': 0, 'total_impressions': 0, 'average_ctr': 0, 'average_position': 0}
        if summary_response.get('rows'):
            row = summary_response['rows'][0]
            summary = {
                'total_clicks': row['clicks'],
                'total_impressions': row['impressions'],
                'average_ctr': round(row['ctr'] * 100, 2),
                'average_position': round(row['position'], 1)
            }

        detail_response = service.searchanalytics().query(
            siteUrl=site_url,
            body={'startDate': start_date, 'endDate': end_date, 'dimensions': ['query'], 'rowLimit': limit}
        ).execute()

        queries = []
        for row in detail_response.get('rows', []):
            queries.append({
                'query': row['keys'][0],
                'clicks': row['clicks'],
                'impressions': row['impressions'],
                'ctr': round(row['ctr'] * 100, 2),
                'position': round(row['position'], 1)
            })

        return jsonify({
            "success": True, "site_url": site_url,
            "start_date": start_date, "end_date": end_date,
            "summary": summary, "query_count": len(queries), "queries": queries
        })

    except Exception as e:
        error_message = str(e)
        if "invalid_grant" in error_message or "401" in error_message:
            return jsonify({"success": False, "error": "トークンが無効です", "error_type": "TOKEN_EXPIRED"}), 401
        return jsonify({"success": False, "error": error_message}), 500

@app.route('/gsc/pages')
def get_gsc_pages():
    try:
        site_url = request.args.get('site_url')
        start_date = request.args.get('start_date', '2025-01-01')
        end_date = request.args.get('end_date', '2025-01-31')
        limit = int(request.args.get('limit', 20))

        if not site_url:
            return jsonify({"success": False, "error": "site_url が必要です"}), 400
        if not all([GSC_REFRESH_TOKEN, GSC_CLIENT_ID, GSC_CLIENT_SECRET]):
            return jsonify({"success": False, "error": "GSC環境変数が設定されていません"}), 500

        creds = Credentials(
            token=None, refresh_token=GSC_REFRESH_TOKEN,
            token_uri='https://oauth2.googleapis.com/token',
            client_id=GSC_CLIENT_ID, client_secret=GSC_CLIENT_SECRET,
            scopes=['https://www.googleapis.com/auth/webmasters.readonly']
        )
        service = build('searchconsole', 'v1', credentials=creds)

        summary_response = service.searchanalytics().query(
            siteUrl=site_url,
            body={'startDate': start_date, 'endDate': end_date}
        ).execute()

        summary = {'total_clicks': 0, 'total_impressions': 0, 'average_ctr': 0, 'average_position': 0}
        if summary_response.get('rows'):
            row = summary_response['rows'][0]
            summary = {
                'total_clicks': row['clicks'],
                'total_impressions': row['impressions'],
                'average_ctr': round(row['ctr'] * 100, 2),
                'average_position': round(row['position'], 1)
            }

        detail_response = service.searchanalytics().query(
            siteUrl=site_url,
            body={'startDate': start_date, 'endDate': end_date, 'dimensions': ['page'], 'rowLimit': limit}
        ).execute()

        pages = []
        for row in detail_response.get('rows', []):
            pages.append({
                'page': row['keys'][0],
                'clicks': row['clicks'],
                'impressions': row['impressions'],
                'ctr': round(row['ctr'] * 100, 2),
                'position': round(row['position'], 1)
            })

        return jsonify({
            "success": True, "site_url": site_url,
            "start_date": start_date, "end_date": end_date,
            "summary": summary, "page_count": len(pages), "pages": pages
        })

    except Exception as e:
        error_message = str(e)
        if "invalid_grant" in error_message or "401" in error_message:
            return jsonify({"success": False, "error": "トークンが無効です", "error_type": "TOKEN_EXPIRED"}), 401
        return jsonify({"success": False, "error": error_message}), 500

@app.route('/google-ads/campaigns')
def get_google_ads_campaigns():
    try:
        customer_id = request.args.get('customer_id')
        start_date = request.args.get('start_date', '2025-01-01')
        end_date = request.args.get('end_date', '2025-01-31')

        if not customer_id:
            return jsonify({"success": False, "error": "customer_id が必要です"}), 400
        if not all([GOOGLE_ADS_REFRESH_TOKEN, GOOGLE_ADS_CLIENT_ID, GOOGLE_ADS_CLIENT_SECRET, GOOGLE_ADS_DEVELOPER_TOKEN]):
            return jsonify({"success": False, "error": "Google広告の環境変数が設定されていません"}), 500

        gaql_query = f"""
            SELECT campaign.name, campaign.status,
                metrics.clicks, metrics.impressions,
                metrics.cost_micros, metrics.conversions, metrics.ctr
            FROM campaign
            WHERE segments.date BETWEEN '{start_date}' AND '{end_date}'
            AND campaign.status = 'ENABLED'
            ORDER BY metrics.cost_micros DESC
        """

        data = query_google_ads(customer_id, gaql_query)

        if 'error' in data:
            error_msg = str(data['error'])
            if 'DEVELOPER_TOKEN_NOT_APPROVED' in error_msg:
                return jsonify({"success": False, "error": "開発者トークンが本番承認されていません", "error_type": "TOKEN_NOT_APPROVED"}), 403
            return jsonify({"success": False, "error": error_msg}), 500

        campaigns = []
        total_clicks = total_impressions = total_cost = total_conversions = 0

        for row in data.get('results', []):
            cost = int(row.get('metrics', {}).get('costMicros', 0)) / 1000000
            clicks = int(row.get('metrics', {}).get('clicks', 0))
            impressions = int(row.get('metrics', {}).get('impressions', 0))
            conversions = float(row.get('metrics', {}).get('conversions', 0))
            ctr = float(row.get('metrics', {}).get('ctr', 0))
            cpa = cost / conversions if conversions > 0 else 0

            campaigns.append({
                'campaign_name': row.get('campaign', {}).get('name', ''),
                'clicks': clicks, 'impressions': impressions,
                'cost': round(cost, 2), 'conversions': round(conversions, 2),
                'ctr': round(ctr * 100, 2), 'cpa': round(cpa, 2)
            })
            total_clicks += clicks
            total_impressions += impressions
            total_cost += cost
            total_conversions += conversions

        return jsonify({
            "success": True,
            "customer_id": customer_id,
            "start_date": start_date,
            "end_date": end_date,
            "summary": {
                'total_clicks': total_clicks,
                'total_impressions': total_impressions,
                'total_cost': round(total_cost, 2),
                'total_conversions': round(total_conversions, 2),
                'average_ctr': round((total_clicks / total_impressions * 100) if total_impressions > 0 else 0, 2),
                'average_cpa': round((total_cost / total_conversions) if total_conversions > 0 else 0, 2)
            },
            "campaign_count": len(campaigns),
            "campaigns": campaigns
        })

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/google-ads/keywords')
def get_google_ads_keywords():
    try:
        customer_id = request.args.get('customer_id')
        start_date = request.args.get('start_date', '2025-01-01')
        end_date = request.args.get('end_date', '2025-01-31')
        limit = int(request.args.get('limit', 20))

        if not customer_id:
            return jsonify({"success": False, "error": "customer_id が必要です"}), 400
        if not all([GOOGLE_ADS_REFRESH_TOKEN, GOOGLE_ADS_CLIENT_ID, GOOGLE_ADS_CLIENT_SECRET, GOOGLE_ADS_DEVELOPER_TOKEN]):
            return jsonify({"success": False, "error": "Google広告の環境変数が設定されていません"}), 500

        gaql_query = f"""
            SELECT ad_group_criterion.keyword.text, ad_group_criterion.keyword.match_type,
                metrics.clicks, metrics.impressions,
                metrics.cost_micros, metrics.conversions, metrics.ctr
            FROM keyword_view
            WHERE segments.date BETWEEN '{start_date}' AND '{end_date}'
            AND ad_group_criterion.status = 'ENABLED'
            ORDER BY metrics.clicks DESC
            LIMIT {limit}
        """

        data = query_google_ads(customer_id, gaql_query)

        if 'error' in data:
            error_msg = str(data['error'])
            if 'DEVELOPER_TOKEN_NOT_APPROVED' in error_msg:
                return jsonify({"success": False, "error": "開発者トークンが本番承認されていません", "error_type": "TOKEN_NOT_APPROVED"}), 403
            return jsonify({"success": False, "error": error_msg}), 500

        keywords = []
        for row in data.get('results', []):
            cost = int(row.get('metrics', {}).get('costMicros', 0)) / 1000000
            conversions = float(row.get('metrics', {}).get('conversions', 0))
            keywords.append({
                'keyword': row.get('adGroupCriterion', {}).get('keyword', {}).get('text', ''),
                'match_type': row.get('adGroupCriterion', {}).get('keyword', {}).get('matchType', ''),
                'clicks': int(row.get('metrics', {}).get('clicks', 0)),
                'impressions': int(row.get('metrics', {}).get('impressions', 0)),
                'cost': round(cost, 2),
                'conversions': round(conversions, 2),
                'ctr': round(float(row.get('metrics', {}).get('ctr', 0)) * 100, 2)
            })

        return jsonify({
            "success": True,
            "customer_id": customer_id,
            "start_date": start_date,
            "end_date": end_date,
            "keyword_count": len(keywords),
            "keywords": keywords
        })

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
