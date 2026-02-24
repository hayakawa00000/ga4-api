from flask import Flask, jsonify, request
from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import RunReportRequest
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.ads.googleads.client import GoogleAdsClient
import os
import json
from datetime import datetime

app = Flask(__name__)

# 環境変数から設定を取得
SERVICE_ACCOUNT_JSON = os.environ.get('SERVICE_ACCOUNT_JSON')
DEFAULT_PROPERTY_ID = os.environ.get('GA4_PROPERTY_ID')

# サーチコンソール用
GSC_REFRESH_TOKEN = os.environ.get('GSC_REFRESH_TOKEN')
GSC_CLIENT_ID = os.environ.get('GSC_CLIENT_ID')
GSC_CLIENT_SECRET = os.environ.get('GSC_CLIENT_SECRET')

# Google広告用
GOOGLE_ADS_REFRESH_TOKEN = os.environ.get('GOOGLE_ADS_REFRESH_TOKEN')
GOOGLE_ADS_CLIENT_ID = os.environ.get('GOOGLE_ADS_CLIENT_ID')
GOOGLE_ADS_CLIENT_SECRET = os.environ.get('GOOGLE_ADS_CLIENT_SECRET')
GOOGLE_ADS_DEVELOPER_TOKEN = os.environ.get('GOOGLE_ADS_DEVELOPER_TOKEN')

def get_google_ads_client():
    """Google広告クライアントを作成"""
    credentials = Credentials(
        token=None,
        refresh_token=GOOGLE_ADS_REFRESH_TOKEN,
        token_uri='https://oauth2.googleapis.com/token',
        client_id=GOOGLE_ADS_CLIENT_ID,
        client_secret=GOOGLE_ADS_CLIENT_SECRET,
        scopes=['https://www.googleapis.com/auth/adwords']
    )
    
    return GoogleAdsClient(
        credentials=credentials,
        developer_token=GOOGLE_ADS_DEVELOPER_TOKEN,
        use_proto_plus=True
    )

@app.route('/')
def home():
    return jsonify({
        "status": "GA4 & GSC & Google Ads API is running",
        "endpoints": {
            "/ga4/sessions": "GA4セッションデータ",
            "/ga4/comprehensive": "GA4包括データ",
            "/gsc/queries": "サーチコンソール検索クエリ",
            "/gsc/pages": "サーチコンソールページ別",
            "/google-ads/campaigns": "Google広告キャンペーン別",
            "/google-ads/keywords": "Google広告キーワード別",
            "/health": "ヘルスチェック"
        }
    })

@app.route('/health')
def health():
    return jsonify({"status": "ok"})

@app.route('/ga4/sessions')
def get_sessions():
    try:
        property_id = request.args.get('property_id', DEFAULT_PROPERTY_ID)
        start_date = request.args.get('start_date', '7daysAgo')
        end_date = request.args.get('end_date', 'today')
        
        if not property_id:
            return jsonify({"success": False, "error": "GA4_PROPERTY_ID 環境変数が設定されていません"}), 500
        
        if start_date not in ['7daysAgo', '30daysAgo', 'yesterday', 'today']:
            try:
                datetime.strptime(start_date, '%Y-%m-%d')
            except ValueError:
                return jsonify({"success": False, "error": "start_date は YYYY-MM-DD 形式で指定してください"}), 400
        
        if end_date not in ['today', 'yesterday']:
            try:
                datetime.strptime(end_date, '%Y-%m-%d')
            except ValueError:
                return jsonify({"success": False, "error": "end_date は YYYY-MM-DD 形式で指定してください"}), 400
        
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
        
        return jsonify({
            "success": True,
            "property_id": property_id,
            "sessions": int(sessions),
            "start_date": start_date,
            "end_date": end_date,
            "message": f"期間 {start_date} 〜 {end_date} のセッション数: {sessions}"
        })
    
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route('/ga4/comprehensive')
def get_comprehensive():
    try:
        property_id = request.args.get('property_id', DEFAULT_PROPERTY_ID)
        start_date = request.args.get('start_date', '7daysAgo')
        end_date = request.args.get('end_date', 'today')
        
        if not property_id:
            return jsonify({"success": False, "error": "GA4_PROPERTY_ID 環境変数が設定されていません"}), 500
        
        if not SERVICE_ACCOUNT_JSON:
            return jsonify({"success": False, "error": "SERVICE_ACCOUNT_JSON が設定されていません"}), 500
        
        service_account_info = json.loads(SERVICE_ACCOUNT_JSON)
        credentials = service_account.Credentials.from_service_account_info(
            service_account_info,
            scopes=['https://www.googleapis.com/auth/analytics.readonly']
        )
        client = BetaAnalyticsDataClient(credentials=credentials)
        
        # 全体サマリー
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
        
        # 参照元/メディア別
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
        
        # デバイス別
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
        
        # ページパス別
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
        
        # 市区町村別
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
        
        # ランディングページ別
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
        
        # イベント名別
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
            "events": events,
            "message": f"期間 {start_date} 〜 {end_date} の包括的なGA4データ"
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
            return jsonify({"success": False, "error": "site_url パラメータが必要です"}), 400
        
        if not all([GSC_REFRESH_TOKEN, GSC_CLIENT_ID, GSC_CLIENT_SECRET]):
            return jsonify({"success": False, "error": "サーチコンソールの環境変数が設定されていません"}), 500
        
        creds = Credentials(
            token=None, refresh_token=GSC_REFRESH_TOKEN,
            token_uri='https://oauth2.googleapis.com/token',
            client_id=GSC_CLIENT_ID, client_secret=GSC_CLIENT_SECRET,
            scopes=['https://www.googleapis.com/auth/webmasters.readonly']
        )
        service = build('searchconsole', 'v1', credentials=creds)
        
        # 全体サマリー
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
        
        # クエリ別詳細
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
            return jsonify({"success": False, "error": "トークンが無効です。再認証が必要です。", "error_type": "TOKEN_EXPIRED"}), 401
        return jsonify({"success": False, "error": error_message}), 500


@app.route('/gsc/pages')
def get_gsc_pages():
    try:
        site_url = request.args.get('site_url')
        start_date = request.args.get('start_date', '2025-01-01')
        end_date = request.args.get('end_date', '2025-01-31')
        limit = int(request.args.get('limit', 20))
        
        if not site_url:
            return jsonify({"success": False, "error": "site_url パラメータが必要です"}), 400
        
        if not all([GSC_REFRESH_TOKEN, GSC_CLIENT_ID, GSC_CLIENT_SECRET]):
            return jsonify({"success": False, "error": "サーチコンソールの環境変数が設定されていません"}), 500
        
        creds = Credentials(
            token=None, refresh_token=GSC_REFRESH_TOKEN,
            token_uri='https://oauth2.googleapis.com/token',
            client_id=GSC_CLIENT_ID, client_secret=GSC_CLIENT_SECRET,
            scopes=['https://www.googleapis.com/auth/webmasters.readonly']
        )
        service = build('searchconsole', 'v1', credentials=creds)
        
        # 全体サマリー
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
        
        # ページ別詳細
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
            return jsonify({"success": False, "error": "トークンが無効です。再認証が必要です。", "error_type": "TOKEN_EXPIRED"}), 401
        return jsonify({"success": False, "error": error_message}), 500


@app.route('/google-ads/campaigns')
def get_google_ads_campaigns():
    try:
        customer_id = request.args.get('customer_id')
        start_date = request.args.get('start_date', '2025-01-01')
        end_date = request.args.get('end_date', '2025-01-31')
        
        if not customer_id:
            return jsonify({"success": False, "error": "customer_id パラメータが必要です（ハイフンなしの数字）"}), 400
        
        if not all([GOOGLE_ADS_REFRESH_TOKEN, GOOGLE_ADS_CLIENT_ID, GOOGLE_ADS_CLIENT_SECRET, GOOGLE_ADS_DEVELOPER_TOKEN]):
            return jsonify({"success": False, "error": "Google広告の環境変数が設定されていません"}), 500
        
        client = get_google_ads_client()
        ga_service = client.get_service("GoogleAdsService")
        
        query = f"""
            SELECT
                campaign.name,
                campaign.status,
                metrics.clicks,
                metrics.impressions,
                metrics.cost_micros,
                metrics.conversions,
                metrics.ctr
            FROM campaign
            WHERE segments.date BETWEEN '{start_date}' AND '{end_date}'
            AND campaign.status = 'ENABLED'
            ORDER BY metrics.cost_micros DESC
        """
        
        response = ga_service.search(customer_id=customer_id, query=query)
        
        campaigns = []
        total_clicks = 0
        total_impressions = 0
        total_cost = 0
        total_conversions = 0
        
        for row in response:
            cost = row.metrics.cost_micros / 1000000
            conversions = row.metrics.conversions
            cpa = cost / conversions if conversions > 0 else 0
            
            campaigns.append({
                'campaign_name': row.campaign.name,
                'clicks': row.metrics.clicks,
                'impressions': row.metrics.impressions,
                'cost': round(cost, 2),
                'conversions': round(conversions, 2),
                'ctr': round(row.metrics.ctr * 100, 2),
                'cpa': round(cpa, 2)
            })
            
            total_clicks += row.metrics.clicks
            total_impressions += row.metrics.impressions
            total_cost += cost
            total_conversions += conversions
        
        summary = {
            'total_clicks': total_clicks,
            'total_impressions': total_impressions,
            'total_cost': round(total_cost, 2),
            'total_conversions': round(total_conversions, 2),
            'average_ctr': round((total_clicks / total_impressions * 100) if total_impressions > 0 else 0, 2),
            'average_cpa': round((total_cost / total_conversions) if total_conversions > 0 else 0, 2)
        }
        
        return jsonify({
            "success": True,
            "customer_id": customer_id,
            "start_date": start_date,
            "end_date": end_date,
            "summary": summary,
            "campaign_count": len(campaigns),
            "campaigns": campaigns,
            "message": f"期間 {start_date} 〜 {end_date} のキャンペーン別データ"
        })
    
    except Exception as e:
        error_message = str(e)
        if "DEVELOPER_TOKEN_NOT_APPROVED" in error_message:
            return jsonify({
                "success": False,
                "error": "開発者トークンが本番承認されていません（テストアカウントのみアクセス可能）",
                "error_type": "TOKEN_NOT_APPROVED"
            }), 403
        return jsonify({"success": False, "error": error_message}), 500


@app.route('/google-ads/keywords')
def get_google_ads_keywords():
    try:
        customer_id = request.args.get('customer_id')
        start_date = request.args.get('start_date', '2025-01-01')
        end_date = request.args.get('end_date', '2025-01-31')
        limit = int(request.args.get('limit', 20))
        
        if not customer_id:
            return jsonify({"success": False, "error": "customer_id パラメータが必要です"}), 400
        
        if not all([GOOGLE_ADS_REFRESH_TOKEN, GOOGLE_ADS_CLIENT_ID, GOOGLE_ADS_CLIENT_SECRET, GOOGLE_ADS_DEVELOPER_TOKEN]):
            return jsonify({"success": False, "error": "Google広告の環境変数が設定されていません"}), 500
        
        client = get_google_ads_client()
        ga_service = client.get_service("GoogleAdsService")
        
        query = f"""
            SELECT
                ad_group_criterion.keyword.text,
                ad_group_criterion.keyword.match_type,
                metrics.clicks,
                metrics.impressions,
                metrics.cost_micros,
                metrics.conversions,
                metrics.ctr
            FROM keyword_view
            WHERE segments.date BETWEEN '{start_date}' AND '{end_date}'
            AND ad_group_criterion.status = 'ENABLED'
            ORDER BY metrics.clicks DESC
            LIMIT {limit}
        """
        
        response = ga_service.search(customer_id=customer_id, query=query)
        
        keywords = []
        for row in response:
            cost = row.metrics.cost_micros / 1000000
            conversions = row.metrics.conversions
            
            keywords.append({
                'keyword': row.ad_group_criterion.keyword.text,
                'match_type': str(row.ad_group_criterion.keyword.match_type.name),
                'clicks': row.metrics.clicks,
                'impressions': row.metrics.impressions,
                'cost': round(cost, 2),
                'conversions': round(conversions, 2),
                'ctr': round(row.metrics.ctr * 100, 2)
            })
        
        return jsonify({
            "success": True,
            "customer_id": customer_id,
            "start_date": start_date,
            "end_date": end_date,
            "keyword_count": len(keywords),
            "keywords": keywords,
            "message": f"期間 {start_date} 〜 {end_date} のキーワード別データ（上位{len(keywords)}件）"
        })
    
    except Exception as e:
        error_message = str(e)
        if "DEVELOPER_TOKEN_NOT_APPROVED" in error_message:
            return jsonify({
                "success": False,
                "error": "開発者トークンが本番承認されていません",
                "error_type": "TOKEN_NOT_APPROVED"
            }), 403
        return jsonify({"success": False, "error": error_message}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
