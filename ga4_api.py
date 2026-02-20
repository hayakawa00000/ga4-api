from flask import Flask, jsonify, request
from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import RunReportRequest
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
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

@app.route('/')
def home():
    return jsonify({
        "status": "GA4 & GSC API is running",
        "default_property_id": DEFAULT_PROPERTY_ID,
        "endpoints": {
            "/ga4/sessions": "Get GA4 sessions data only",
            "/ga4/comprehensive": "Get comprehensive GA4 data",
            "/gsc/queries": "Get Search Console queries data with summary",
            "/gsc/pages": "Get Search Console pages data with summary",
            "/health": "Health check"
        },
        "usage": {
            "ga4_sessions": "/ga4/sessions?start_date=2024-01-01&end_date=2024-01-31",
            "ga4_comprehensive": "/ga4/comprehensive?start_date=2024-01-01&end_date=2024-01-31",
            "gsc_queries": "/gsc/queries?site_url=https://your-site.com&start_date=2024-01-01&end_date=2024-01-31",
            "gsc_pages": "/gsc/pages?site_url=https://your-site.com&start_date=2024-01-01&end_date=2024-01-31"
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
            return jsonify({
                "success": False,
                "error": "GA4_PROPERTY_ID 環境変数が設定されていません"
            }), 500
        
        if start_date not in ['7daysAgo', '30daysAgo', 'yesterday', 'today']:
            try:
                datetime.strptime(start_date, '%Y-%m-%d')
            except ValueError:
                return jsonify({
                    "success": False,
                    "error": "start_date は YYYY-MM-DD 形式で指定してください（例: 2024-01-01）"
                }), 400
        
        if end_date not in ['today', 'yesterday']:
            try:
                datetime.strptime(end_date, '%Y-%m-%d')
            except ValueError:
                return jsonify({
                    "success": False,
                    "error": "end_date は YYYY-MM-DD 形式で指定してください（例: 2024-01-31）"
                }), 400
        
        if not SERVICE_ACCOUNT_JSON:
            return jsonify({
                "success": False,
                "error": "SERVICE_ACCOUNT_JSON が設定されていません"
            }), 500
        
        service_account_info = json.loads(SERVICE_ACCOUNT_JSON)
        credentials = service_account.Credentials.from_service_account_info(
            service_account_info,
            scopes=['https://www.googleapis.com/auth/analytics.readonly']
        )
        
        client = BetaAnalyticsDataClient(credentials=credentials)
        
        request_obj = RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[{
                "start_date": start_date,
                "end_date": end_date
            }],
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
    
    except json.JSONDecodeError as e:
        return jsonify({
            "success": False,
            "error": "SERVICE_ACCOUNT_JSON の形式が正しくありません",
            "details": str(e)
        }), 500
    
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


@app.route('/ga4/comprehensive')
def get_comprehensive():
    try:
        property_id = request.args.get('property_id', DEFAULT_PROPERTY_ID)
        start_date = request.args.get('start_date', '7daysAgo')
        end_date = request.args.get('end_date', 'today')
        
        if not property_id:
            return jsonify({
                "success": False,
                "error": "GA4_PROPERTY_ID 環境変数が設定されていません"
            }), 500
        
        if start_date not in ['7daysAgo', '30daysAgo', 'yesterday', 'today']:
            try:
                datetime.strptime(start_date, '%Y-%m-%d')
            except ValueError:
                return jsonify({
                    "success": False,
                    "error": "start_date は YYYY-MM-DD 形式で指定してください"
                }), 400
        
        if end_date not in ['today', 'yesterday']:
            try:
                datetime.strptime(end_date, '%Y-%m-%d')
            except ValueError:
                return jsonify({
                    "success": False,
                    "error": "end_date は YYYY-MM-DD 形式で指定してください"
                }), 400
        
        if not SERVICE_ACCOUNT_JSON:
            return jsonify({
                "success": False,
                "error": "SERVICE_ACCOUNT_JSON が設定されていません"
            }), 500
        
        service_account_info = json.loads(SERVICE_ACCOUNT_JSON)
        credentials = service_account.Credentials.from_service_account_info(
            service_account_info,
            scopes=['https://www.googleapis.com/auth/analytics.readonly']
        )
        
        client = BetaAnalyticsDataClient(credentials=credentials)
        
        # 全体サマリー
        summary_request = RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[{"start_date": start_date, "end_date": end_date}],
            metrics=[
                {"name": "sessions"},
                {"name": "activeUsers"},
                {"name": "screenPageViews"},
                {"name": "engagementRate"},
                {"name": "bounceRate"},
                {"name": "averageSessionDuration"},
                {"name": "screenPageViewsPerSession"},
                {"name": "keyEvents"}
            ]
        )
        summary_response = client.run_report(summary_request)
        
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
        else:
            summary = {}
        
        # 参照元/メディア別
        source_request = RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[{"start_date": start_date, "end_date": end_date}],
            dimensions=[
                {"name": "sessionSource"},
                {"name": "sessionMedium"}
            ],
            metrics=[
                {"name": "sessions"},
                {"name": "activeUsers"}
            ],
            order_bys=[{"metric": {"metric_name": "sessions"}, "desc": True}],
            limit=10
        )
        source_response = client.run_report(source_request)
        
        traffic_sources = []
        for row in source_response.rows:
            traffic_sources.append({
                "source": row.dimension_values[0].value,
                "medium": row.dimension_values[1].value,
                "sessions": int(row.metric_values[0].value),
                "users": int(row.metric_values[1].value)
            })
        
        # デバイス別
        device_request = RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[{"start_date": start_date, "end_date": end_date}],
            dimensions=[{"name": "deviceCategory"}],
            metrics=[
                {"name": "sessions"},
                {"name": "activeUsers"},
                {"name": "engagementRate"}
            ],
            order_bys=[{"metric": {"metric_name": "sessions"}, "desc": True}]
        )
        device_response = client.run_report(device_request)
        
        devices = []
        for row in device_response.rows:
            devices.append({
                "device": row.dimension_values[0].value,
                "sessions": int(row.metric_values[0].value),
                "users": int(row.metric_values[1].value),
                "engagement_rate": float(row.metric_values[2].value)
            })
        
        # ページパス別
        page_request = RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[{"start_date": start_date, "end_date": end_date}],
            dimensions=[{"name": "pagePath"}],
            metrics=[
                {"name": "screenPageViews"},
                {"name": "activeUsers"},
                {"name": "averageSessionDuration"}
            ],
            order_bys=[{"metric": {"metric_name": "screenPageViews"}, "desc": True}],
            limit=20
        )
        page_response = client.run_report(page_request)
        
        pages = []
        for row in page_response.rows:
            pages.append({
                "page_path": row.dimension_values[0].value,
                "pageviews": int(row.metric_values[0].value),
                "users": int(row.metric_values[1].value),
                "avg_session_duration": float(row.metric_values[2].value)
            })
        
        # 市区町村別
        city_request = RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[{"start_date": start_date, "end_date": end_date}],
            dimensions=[{"name": "city"}],
            metrics=[
                {"name": "sessions"},
                {"name": "activeUsers"}
            ],
            order_bys=[{"metric": {"metric_name": "sessions"}, "desc": True}],
            limit=10
        )
        city_response = client.run_report(city_request)
        
        cities = []
        for row in city_response.rows:
            cities.append({
                "city": row.dimension_values[0].value,
                "sessions": int(row.metric_values[0].value),
                "users": int(row.metric_values[1].value)
            })
        
        # ランディングページ別
        landing_request = RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[{"start_date": start_date, "end_date": end_date}],
            dimensions=[{"name": "landingPage"}],
            metrics=[
                {"name": "sessions"},
                {"name": "bounceRate"},
                {"name": "engagementRate"}
            ],
            order_bys=[{"metric": {"metric_name": "sessions"}, "desc": True}],
            limit=10
        )
        landing_response = client.run_report(landing_request)
        
        landing_pages = []
        for row in landing_response.rows:
            landing_pages.append({
                "landing_page": row.dimension_values[0].value,
                "sessions": int(row.metric_values[0].value),
                "bounce_rate": float(row.metric_values[1].value),
                "engagement_rate": float(row.metric_values[2].value)
            })
        
        # イベント名別
        event_request = RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[{"start_date": start_date, "end_date": end_date}],
            dimensions=[{"name": "eventName"}],
            metrics=[{"name": "eventCount"}],
            order_bys=[{"metric": {"metric_name": "eventCount"}, "desc": True}],
            limit=10
        )
        event_response = client.run_report(event_request)
        
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
    
    except json.JSONDecodeError as e:
        return jsonify({
            "success": False,
            "error": "SERVICE_ACCOUNT_JSON の形式が正しくありません",
            "details": str(e)
        }), 500
    
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


@app.route('/gsc/queries')
def get_gsc_queries():
    try:
        site_url = request.args.get('site_url')
        start_date = request.args.get('start_date', '2025-01-01')
        end_date = request.args.get('end_date', '2025-01-31')
        limit = int(request.args.get('limit', 20))
        
        if not site_url:
            return jsonify({
                "success": False,
                "error": "site_url パラメータが必要です",
                "usage": "/gsc/queries?site_url=https://your-site.com&start_date=2025-01-01&end_date=2025-01-31"
            }), 400
        
        if not all([GSC_REFRESH_TOKEN, GSC_CLIENT_ID, GSC_CLIENT_SECRET]):
            return jsonify({
                "success": False,
                "error": "サーチコンソールの環境変数が設定されていません"
            }), 500
        
        # OAuth認証
        creds = Credentials(
            token=None,
            refresh_token=GSC_REFRESH_TOKEN,
            token_uri='https://oauth2.googleapis.com/token',
            client_id=GSC_CLIENT_ID,
            client_secret=GSC_CLIENT_SECRET,
            scopes=['https://www.googleapis.com/auth/webmasters.readonly']
        )
        
        # Search Console API
        service = build('searchconsole', 'v1', credentials=creds)
        
        # 1. 全体サマリーを取得（ディメンションなし）
        summary_request = {
            'startDate': start_date,
            'endDate': end_date
        }
        
        summary_response = service.searchanalytics().query(
            siteUrl=site_url,
            body=summary_request
        ).execute()
        
        # 全体の合計値を取得
        if summary_response.get('rows'):
            summary_row = summary_response['rows'][0]
            summary = {
                'total_clicks': summary_row['clicks'],
                'total_impressions': summary_row['impressions'],
                'average_ctr': round(summary_row['ctr'] * 100, 2),
                'average_position': round(summary_row['position'], 1)
            }
        else:
            summary = {
                'total_clicks': 0,
                'total_impressions': 0,
                'average_ctr': 0,
                'average_position': 0
            }
        
        # 2. クエリ別の詳細データを取得
        detail_request = {
            'startDate': start_date,
            'endDate': end_date,
            'dimensions': ['query'],
            'rowLimit': limit
        }
        
        detail_response = service.searchanalytics().query(
            siteUrl=site_url,
            body=detail_request
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
            "success": True,
            "site_url": site_url,
            "start_date": start_date,
            "end_date": end_date,
            "summary": summary,
            "query_count": len(queries),
            "queries": queries,
            "message": f"{site_url} の検索データ（全体サマリー + 上位{len(queries)}件の詳細）"
        })
    
    except Exception as e:
        error_message = str(e)
        
        if "invalid_grant" in error_message or "401" in error_message:
            return jsonify({
                "success": False,
                "error": "トークンが無効です。再認証が必要です。",
                "error_type": "TOKEN_EXPIRED",
                "details": error_message
            }), 401
        else:
            return jsonify({
                "success": False,
                "error": error_message
            }), 500


@app.route('/gsc/pages')
def get_gsc_pages():
    try:
        site_url = request.args.get('site_url')
        start_date = request.args.get('start_date', '2025-01-01')
        end_date = request.args.get('end_date', '2025-01-31')
        limit = int(request.args.get('limit', 20))
        
        if not site_url:
            return jsonify({
                "success": False,
                "error": "site_url パラメータが必要です",
                "usage": "/gsc/pages?site_url=https://your-site.com&start_date=2025-01-01&end_date=2025-01-31"
            }), 400
        
        if not all([GSC_REFRESH_TOKEN, GSC_CLIENT_ID, GSC_CLIENT_SECRET]):
            return jsonify({
                "success": False,
                "error": "サーチコンソールの環境変数が設定されていません"
            }), 500
        
        # OAuth認証
        creds = Credentials(
            token=None,
            refresh_token=GSC_REFRESH_TOKEN,
            token_uri='https://oauth2.googleapis.com/token',
            client_id=GSC_CLIENT_ID,
            client_secret=GSC_CLIENT_SECRET,
            scopes=['https://www.googleapis.com/auth/webmasters.readonly']
        )
        
        # Search Console API
        service = build('searchconsole', 'v1', credentials=creds)
        
        # 1. 全体サマリーを取得
        summary_request = {
            'startDate': start_date,
            'endDate': end_date
        }
        
        summary_response = service.searchanalytics().query(
            siteUrl=site_url,
            body=summary_request
        ).execute()
        
        if summary_response.get('rows'):
            summary_row = summary_response['rows'][0]
            summary = {
                'total_clicks': summary_row['clicks'],
                'total_impressions': summary_row['impressions'],
                'average_ctr': round(summary_row['ctr'] * 100, 2),
                'average_position': round(summary_row['position'], 1)
            }
        else:
            summary = {
                'total_clicks': 0,
                'total_impressions': 0,
                'average_ctr': 0,
                'average_position': 0
            }
        
        # 2. ページ別の詳細データを取得
        detail_request = {
            'startDate': start_date,
            'endDate': end_date,
            'dimensions': ['page'],
            'rowLimit': limit
        }
        
        detail_response = service.searchanalytics().query(
            siteUrl=site_url,
            body=detail_request
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
            "success": True,
            "site_url": site_url,
            "start_date": start_date,
            "end_date": end_date,
            "summary": summary,
            "page_count": len(pages),
            "pages": pages,
            "message": f"{site_url} のページ別データ（全体サマリー + 上位{len(pages)}件の詳細）"
        })
    
    except Exception as e:
        error_message = str(e)
        
        if "invalid_grant" in error_message or "401" in error_message:
            return jsonify({
                "success": False,
                "error": "トークンが無効です。再認証が必要です。",
                "error_type": "TOKEN_EXPIRED",
                "details": error_message
            }), 401
        else:
            return jsonify({
                "success": False,
                "error": error_message
            }), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
