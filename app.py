from flask import Flask, request, jsonify
import pandas as pd
import os
import requests
from io import BytesIO
from openai import OpenAI

app = Flask(__name__)

# ============================================
# CONFIGURATION - Uses Environment Variables
# ============================================
EXCEL_FILE = 'actuarial_sophisticated_sample.xlsx'
LOSS_RATIO_THRESHOLD = 75.0

# Read from environment variables (set in Render dashboard)
SLACK_BOT_TOKEN = os.environ.get('SLACK_BOT_TOKEN', 'YOUR_SLACK_BOT_TOKEN_HERE')
DEEPSEEK_API_KEY = os.environ.get('DEEPSEEK_API_KEY', 'YOUR_DEEPSEEK_API_KEY_HERE')
# ============================================

# Track the most recently uploaded file and analysis
LAST_UPLOADED_FILE = None
LAST_UPLOADED_FILE_NAME = None
LAST_ANALYSIS_RESULT = None

# Initialize DeepSeek client
deepseek_client = OpenAI(
    api_key=DEEPSEEK_API_KEY,
    base_url="https://api.deepseek.com"
)

def calculate_loss_ratio(file_source=None):
    """
    Read Excel file and calculate actuarial loss ratio
    Args:
        file_source: Either a file path (str) or file URL (str) or None (uses default)
    Returns: dict with premium, claims, and loss ratio
    """
    try:
        # Read Excel file from different sources
        if file_source and file_source.startswith('http'):
            # Download file from URL (for Slack file uploads)
            headers = {'Authorization': f'Bearer {SLACK_BOT_TOKEN}'}
            response = requests.get(file_source, headers=headers)
            response.raise_for_status()
            df = pd.read_excel(BytesIO(response.content))
        elif file_source:
            # Read from local file path
            df = pd.read_excel(file_source)
        else:
            # Use default file
            df = pd.read_excel(EXCEL_FILE)
        
        # Validate required columns exist
        required_columns = ['Premium', 'Claims']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            return {
                'success': False, 
                'error': f'Missing required columns: {", ".join(missing_columns)}. File must have "Premium" and "Claims" columns.'
            }
        
        # Calculate totals
        total_premium = df['Premium'].sum()
        total_claims = df['Claims'].sum()
        
        # Calculate loss ratio
        loss_ratio = (total_claims / total_premium) * 100 if total_premium > 0 else 0
        
        return {
            'success': True,
            'premium': total_premium,
            'claims': total_claims,
            'loss_ratio': loss_ratio,
            'num_policies': len(df)
        }
    
    except FileNotFoundError:
        return {'success': False, 'error': 'Excel file not found'}
    except KeyError as e:
        return {'success': False, 'error': f'Missing column in Excel: {e}'}
    except requests.exceptions.RequestException as e:
        return {'success': False, 'error': f'Failed to download file: {str(e)}'}
    except Exception as e:
        return {'success': False, 'error': f'Error reading file: {str(e)}'}

def generate_ai_insights(result):
    """
    Generate AI-powered insights about the loss ratio analysis
    """
    if not result['success']:
        return None
    
    try:
        # Create prompt for DeepSeek
        prompt = f"""You are an expert actuary analyzing insurance portfolio data.

Analysis Results:
- Total Premium: ${result['premium']:,.0f}
- Total Claims: ${result['claims']:,.0f}
- Loss Ratio: {result['loss_ratio']:.1f}%
- Number of Policies: {result['num_policies']}
- Risk Threshold: {LOSS_RATIO_THRESHOLD}%

Provide a brief, professional 2-3 sentence insight about:
1. What this loss ratio indicates about portfolio health
2. Whether it's above/below threshold and what that means
3. One actionable recommendation

Keep it concise and business-focused."""

        # Call DeepSeek API
        response = deepseek_client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "You are an expert actuarial analyst providing concise, actionable insights."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=200
        )
        
        insight = response.choices[0].message.content.strip()
        return insight
    
    except Exception as e:
        print(f"❌ AI insight generation failed: {e}")
        return None

def answer_actuarial_question(question, context_result):
    """
    Use AI to answer questions about the actuarial analysis
    """
    if not context_result or not context_result.get('success'):
        return "I don't have any analysis data to reference. Please upload an Excel file or run /lossratio first."
    
    try:
        # Create context-aware prompt
        prompt = f"""You are an expert actuary. A user is asking about their portfolio analysis.

Current Analysis Context:
- Total Premium: ${context_result['premium']:,.0f}
- Total Claims: ${context_result['claims']:,.0f}
- Loss Ratio: {context_result['loss_ratio']:.1f}%
- Number of Policies: {context_result['num_policies']}
- Risk Threshold: {LOSS_RATIO_THRESHOLD}%

User Question: {question}

Provide a clear, professional answer based on the analysis data above. Be specific and reference the actual numbers. Keep it under 4 sentences."""

        # Call DeepSeek API
        response = deepseek_client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "You are an expert actuarial consultant providing clear, data-driven answers."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=300
        )
        
        answer = response.choices[0].message.content.strip()
        return answer
    
    except Exception as e:
        print(f"❌ AI answer generation failed: {e}")
        return f"Sorry, I couldn't generate an answer. Error: {str(e)}"

def format_slack_response(result, file_name=None, include_ai=True):
    """
    Format the calculation results for Slack with optional AI insights
    """
    if not result['success']:
        file_info = f" ({file_name})" if file_name else ""
        return {
            'response_type': 'in_channel',
            'text': f"❌ *Error{file_info}:* {result['error']}"
        }
    
    # Check if loss ratio exceeds threshold
    warning = ""
    if result['loss_ratio'] > LOSS_RATIO_THRESHOLD:
        warning = f"\n⚠️ *Warning:* Loss ratio exceeds {LOSS_RATIO_THRESHOLD}% threshold!"
    
    # Format currency
    premium_formatted = f"${result['premium']:,.0f}"
    claims_formatted = f"${result['claims']:,.0f}"
    
    # Add file name if provided
    file_header = f"📄 *Analysis of: {file_name}*\n\n" if file_name else ""
    
    # Create basic message
    message = f"""{file_header}📊 *Actuarial Loss Ratio Analysis*

*Premium:* {premium_formatted}
*Claims:* {claims_formatted}
*Loss Ratio:* {result['loss_ratio']:.1f}%
*Policies Analyzed:* {result['num_policies']}{warning}
"""
    
    # Add AI insights if enabled
    if include_ai and DEEPSEEK_API_KEY != 'YOUR_DEEPSEEK_API_KEY_HERE':
        print("🤖 Generating AI insights...")
        ai_insight = generate_ai_insights(result)
        if ai_insight:
            message += f"\n🤖 *AI Insights:*\n_{ai_insight}_"
    
    return {
        'response_type': 'in_channel',
        'text': message
    }

def send_message_to_channel(channel_id, result, file_name=None, include_ai=True):
    """
    Send a formatted message to a Slack channel
    """
    try:
        from slack_sdk import WebClient
        
        client = WebClient(token=SLACK_BOT_TOKEN)
        
        # Format the message
        response = format_slack_response(result, file_name, include_ai)
        
        # Send to channel
        client.chat_postMessage(channel=channel_id, text=response['text'])
        print(f"✅ Sent analysis to channel {channel_id}")
    except Exception as e:
        print(f"❌ Error sending message: {e}")

@app.route('/slack/events', methods=['POST'])
def slack_events():
    """
    Handle Slack events (like file uploads)
    """
    global LAST_UPLOADED_FILE, LAST_UPLOADED_FILE_NAME, LAST_ANALYSIS_RESULT
    
    data = request.json
    
    # Handle URL verification challenge (first time setup)
    if 'challenge' in data:
        print("✅ Received Slack verification challenge")
        return jsonify({'challenge': data['challenge']})
    
    # Ignore retry attempts (Slack sends these if we're slow)
    if request.headers.get('X-Slack-Retry-Num'):
        print("⚠️ Ignoring retry request")
        return jsonify({'status': 'ok'})
    
    # Handle actual events
    event = data.get('event', {})
    event_type = event.get('type')
    
    print(f"📨 Received event: {event_type}")
    
    # Check if it's a file share event
    if event_type == 'message' and 'files' in event:
        # Get the first uploaded file
        file_info = event['files'][0]
        file_url = file_info.get('url_private')
        file_name = file_info.get('name', 'unknown')
        
        print(f"📎 File uploaded: {file_name}")
        
        # Check if it's an Excel file
        if file_name.endswith(('.xlsx', '.xls')):
            # Save as the last uploaded file
            LAST_UPLOADED_FILE = file_url
            LAST_UPLOADED_FILE_NAME = file_name
            print(f"💾 Saved as last uploaded file: {file_name}")
            
            # Analyze the uploaded file
            print(f"🔍 Analyzing {file_name}...")
            result = calculate_loss_ratio(file_url)
            
            # Save the analysis for /explain command
            LAST_ANALYSIS_RESULT = result
            
            # Send response back to the channel (with AI)
            send_message_to_channel(event['channel'], result, file_name, include_ai=True)
        else:
            # Not an Excel file
            print(f"⚠️ File is not Excel: {file_name}")
            error_result = {
                'success': False, 
                'error': f'Please upload an Excel file (.xlsx or .xls). You uploaded: {file_name}'
            }
            send_message_to_channel(event['channel'], error_result, include_ai=False)
    
    return jsonify({'status': 'ok'})

@app.route('/lossratio', methods=['POST'])
def lossratio_command():
    """
    Handle the /lossratio slash command from Slack
    Uses the most recently uploaded file, or falls back to default
    """
    global LAST_UPLOADED_FILE, LAST_UPLOADED_FILE_NAME, LAST_ANALYSIS_RESULT
    
    print("📊 /lossratio command received")
    
    # Decide which file to analyze
    if LAST_UPLOADED_FILE:
        print(f"📂 Using last uploaded file: {LAST_UPLOADED_FILE_NAME}")
        result = calculate_loss_ratio(LAST_UPLOADED_FILE)
        file_name = LAST_UPLOADED_FILE_NAME
    else:
        print(f"📂 Using default file: {EXCEL_FILE}")
        result = calculate_loss_ratio()
        file_name = None
    
    # Save the analysis for /explain command
    LAST_ANALYSIS_RESULT = result
    
    # Format and return response (with AI)
    response = format_slack_response(result, file_name, include_ai=True)
    return jsonify(response)

@app.route('/explain', methods=['POST'])
def explain_command():
    """
    Handle the /explain slash command - AI-powered Q&A about the analysis
    """
    global LAST_ANALYSIS_RESULT
    
    print("🤖 /explain command received")
    
    # Get the user's question and channel info
    question = request.form.get('text', '').strip()
    channel_id = request.form.get('channel_id')
    user_id = request.form.get('user_id')
    
    if not question:
        return jsonify({
            'response_type': 'ephemeral',
            'text': "❓ Please ask a question! Example: `/explain why is the loss ratio high?`"
        })
    
    # Check if we have analysis data
    if not LAST_ANALYSIS_RESULT or not LAST_ANALYSIS_RESULT.get('success'):
        return jsonify({
            'response_type': 'ephemeral',
            'text': "⚠️ No analysis data available. Please upload a file or run `/lossratio` first."
        })
    
    # Immediately respond to avoid timeout (Slack requires response within 3 seconds)
    # We'll post the actual answer separately
    immediate_response = jsonify({
        'response_type': 'in_channel',
        'text': f"❓ *Question:* {question}\n\n🤖 _Thinking..._"
    })
    
    # Generate AI answer in background and post to channel
    import threading
    
    def generate_and_post_answer():
        try:
            from slack_sdk import WebClient
            
            print(f"🤖 Generating answer for: {question}")
            answer = answer_actuarial_question(question, LAST_ANALYSIS_RESULT)
            
            # Post the answer to the channel
            client = WebClient(token=SLACK_BOT_TOKEN)
            message = f"🤖 *AI Answer:*\n{answer}"
            client.chat_postMessage(channel=channel_id, text=message)
            print(f"✅ Posted AI answer to channel")
        except Exception as e:
            print(f"❌ Error posting answer: {e}")
    
    # Start background thread
    thread = threading.Thread(target=generate_and_post_answer)
    thread.start()
    
    return immediate_response

@app.route('/health', methods=['GET'])
def health_check():
    """
    Simple health check endpoint
    """
    slack_status = "✅ Set" if SLACK_BOT_TOKEN != 'YOUR_SLACK_BOT_TOKEN_HERE' else "❌ NOT SET"
    ai_status = "✅ Set" if DEEPSEEK_API_KEY != 'YOUR_DEEPSEEK_API_KEY_HERE' else "❌ NOT SET"
    
    return jsonify({
        'status': 'ok', 
        'message': 'Actuarial Slackbot is running!',
        'slack_token': slack_status,
        'deepseek_api': ai_status,
        'deployed': True
    })

@app.route('/', methods=['GET'])
def home():
    """
    Home page showing bot is running
    """
    return """
    <html>
    <body style="font-family: Arial; padding: 40px; background: #f5f5f5;">
        <h1>🤖 Actuarial Slackbot</h1>
        <p>AI-powered actuarial analysis tool</p>
        <h3>Status: ✅ Running</h3>
        <h3>Features:</h3>
        <ul>
            <li>📊 Automated loss ratio calculations</li>
            <li>📤 Excel file upload support</li>
            <li>🤖 AI-powered insights (DeepSeek)</li>
            <li>💬 Conversational Q&A with /explain</li>
        </ul>
        <p><a href="/health">Check Health Status</a></p>
    </body>
    </html>
    """

if __name__ == '__main__':
    # Get port from environment variable (Render sets this)
    port = int(os.environ.get('PORT', 3000))
    
    # Check configuration
    print("\n" + "="*50)
    print("🤖 ACTUARIAL SLACKBOT - DEPLOYED VERSION")
    print("="*50)
    
    if SLACK_BOT_TOKEN == 'YOUR_SLACK_BOT_TOKEN_HERE':
        print("\n⚠️  WARNING: Slack token not configured!")
    else:
        print("✅ Slack token configured")
    
    if DEEPSEEK_API_KEY == 'YOUR_DEEPSEEK_API_KEY_HERE':
        print("⚠️  WARNING: DeepSeek API key not configured!")
        print("   AI features will be disabled")
    else:
        print("✅ DeepSeek API configured")
    
    print(f"\n🚀 Starting server on port {port}...")
    print("="*50 + "\n")
    
    # Run Flask app
    app.run(host='0.0.0.0', port=port, debug=False)