from flask import Flask, jsonify, request
from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import RunReportRequest
from google.oauth2 import service_account
import os
import json
from datetime import datetime

app = Flask(__name__)

# 環境変数から設定を取得
SERVICE_ACCOUNT_JSON = os.environ.get('SERVICE_ACCOUNT_JSON')
DEFAULT_PROPERTY_ID = os.environ.get('GA4_PROPERTY_ID')

@app.route('/')
def home():
    return jsonify({
        "status": "GA4 API is running",
        "default_property_id": DEFAULT_PROPERTY_ID,
        "endpoints": {
            "/ga4/sessions": "Get sessions data only",
            "/ga4/comprehensive": "Get comprehensive GA4 data",
            "/health": "Health check"
        },
        "usage": {
            "sessions": "/ga4/sessions?start_date=2024-01-01&end_date=2024-01-31",
            "comprehensive": "/ga4/comprehensive?start_date=2024-01-01&end_date=2024-01-31"
        }
    })

@app.route('/health')
def health():
    return jsonify({"status": "ok"})

@app.route('/ga4/sessions')
def get_sessions():
    try:
        # URLパラメータを取得
        property_id = request.args.get('property_id', DEFAULT_PROPERTY_ID)
        start_date = request.args.get('start_date', '7daysAgo')
        end_date = request.args.get('end_date', 'today')
        
        # プロパティIDのチェック
        if not property_id:
            return jsonify({
                "success": False,
                "error": "GA4_PROPERTY_ID 環境変数が設定されていません"
            }), 500
        
        # 日付フォーマットの検証
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
        
        # 環境変数のチェック
        if not SERVICE_ACCOUNT_JSON:
            return jsonify({
                "success": False,
                "error": "SERVICE_ACCOUNT_JSON が設定されていません"
            }), 500
        
        # サービスアカウント情報をパース
        service_account_info = json.loads(SERVICE_ACCOUNT_JSON)
        
        # 認証情報を作成
        credentials = service_account.Credentials.from_service_account_info(
            service_account_info,
            scopes=['https://www.googleapis.com/auth/analytics.readonly']
        )
        
        # クライアント作成
        client = BetaAnalyticsDataClient(credentials=credentials)
        
        # レポートリクエスト
        request_obj = RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[{
                "start_date": start_date,
                "end_date": end_date
            }],
            metrics=[{"name": "sessions"}]
        )
        
        # APIを実行
        response = client.run_report(request_obj)
        
        # セッション数を取得
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
        # URLパラメータを取得
        property_id = request.args.get('property_id', DEFAULT_PROPERTY_ID)
        start_date = request.args.get('start_date', '7daysAgo')
        end_date = request.args.get('end_date', 'today')
        
        # プロパティIDのチェック
        if not property_id:
            return jsonify({
                "success": False,
                "error": "GA4_PROPERTY_ID 環境変数が設定されていません"
            }), 500
        
        # 日付フォーマットの検証
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
        
        # 環境変数のチェック
        if not SERVICE_ACCOUNT_JSON:
            return jsonify({
                "success": False,
                "error": "SERVICE_ACCOUNT_JSON が設定されていません"
            }), 500
        
        # サービスアカウント情報をパース
        service_account_info = json.loads(SERVICE_ACCOUNT_JSON)
        
        # 認証情報を作成
        credentials = service_account.Credentials.from_service_account_info(
            service_account_info,
            scopes=['https://www.googleapis.com/auth/analytics.readonly']
        )
        
        # クライアント作成
        client = BetaAnalyticsDataClient(credentials=credentials)
        
        # === 1. 全体サマリー ===
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
        
        # === 2. 参照元/メディア別 ===
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
        
        # === 3. デバイス別 ===
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
        
        # === 4. ページパス別（上位20） ===
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
        
        # === 5. 市区町村別（上位10） ===
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
        
        # === 6. ランディングページ別（上位10） ===
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
        
        # === 7. イベント名別（上位10） ===
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
        
        # === レスポンスをまとめて返す ===
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


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
