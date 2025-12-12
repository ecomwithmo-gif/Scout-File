from flask import Flask, Response
import json

app = Flask(__name__)

# HTML template for the landing page
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Formatter Pro - Scout File</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }
        
        .container {
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            max-width: 800px;
            width: 100%;
            padding: 40px;
            text-align: center;
        }
        
        h1 {
            color: #1a202c;
            font-size: 2.5rem;
            margin-bottom: 10px;
        }
        
        .subtitle {
            color: #718096;
            font-size: 1.2rem;
            margin-bottom: 30px;
        }
        
        .description {
            color: #4a5568;
            line-height: 1.8;
            margin-bottom: 40px;
            text-align: left;
        }
        
        .features {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 40px;
        }
        
        .feature {
            background: #f7fafc;
            padding: 20px;
            border-radius: 10px;
            border: 2px solid #e2e8f0;
        }
        
        .feature h3 {
            color: #2d3748;
            margin-bottom: 10px;
        }
        
        .feature p {
            color: #718096;
            font-size: 0.9rem;
        }
        
        .cta {
            margin-top: 30px;
        }
        
        .btn {
            display: inline-block;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px 40px;
            border-radius: 10px;
            text-decoration: none;
            font-weight: bold;
            font-size: 1.1rem;
            transition: transform 0.2s, box-shadow 0.2s;
            box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
        }
        
        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(102, 126, 234, 0.6);
        }
        
        .github-link {
            margin-top: 20px;
            display: inline-block;
            color: #667eea;
            text-decoration: none;
        }
        
        .github-link:hover {
            text-decoration: underline;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Excel Formatter Pro</h1>
        <p class="subtitle">Advanced Excel Processing Application</p>
        
        <div class="description">
            <p>Excel Formatter Pro is a powerful desktop application for processing and formatting Excel files with advanced analytics, streaming support, and a beautiful modern interface.</p>
        </div>
        
        <div class="features">
            <div class="feature">
                <h3>⚡ Fast Processing</h3>
                <p>Handle large Excel files efficiently with intelligent streaming</p>
            </div>
            <div class="feature">
                <h3>📈 Analytics</h3>
                <p>Comprehensive data analysis and visualization</p>
            </div>
            <div class="feature">
                <h3>🎨 Modern UI</h3>
                <p>Clean, responsive interface with dark/light themes</p>
            </div>
        </div>
        
        <div class="cta">
            <a href="https://github.com/ecomwithmo-gif/Scout-File" class="btn" target="_blank">
                View on GitHub
            </a>
            <br>
            <a href="https://github.com/ecomwithmo-gif/Scout-File" class="github-link" target="_blank">
                Get Started →
            </a>
        </div>
    </div>
</body>
</html>
"""

@app.route('/')
def index():
    return HTML_TEMPLATE

@app.route('/api/health')
def health():
    return Response(
        json.dumps({'status': 'ok', 'message': 'Excel Formatter Pro API is running'}),
        mimetype='application/json'
    )
