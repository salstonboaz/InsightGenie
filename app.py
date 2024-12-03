import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from flask import Flask, render_template, request, redirect, url_for, send_file, flash
from werkzeug.utils import secure_filename  # Ensure correct import

app = Flask(__name__)
app.secret_key = 'supersecretkey'

# Configure upload folder
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
REPORT_FOLDER = os.path.join(os.getcwd(), 'reports')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['REPORT_FOLDER'] = REPORT_FOLDER

# Ensure the folders exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)

# Home page route
@app.route('/')
def home():
    return render_template('index.html')

# File upload and processing route
@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash('No file selected')
        return redirect(url_for('home'))

    # Save the file to the upload folder
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)

    # Ask user if first two columns consist of names
    columns_as_names = request.form.get('columns_as_names', 'no')

    # Read the file using pandas
    df = pd.read_excel(file_path)

    # Check if the first two columns are names (if indicated by the user)
    if columns_as_names.lower() == 'yes':
        df = df.iloc[:, 2:]  # Exclude the first two columns (names)

    # Perform descriptive statistics
    descriptive_stats = df.describe()
    descriptive_interpretation = generate_interpretation(descriptive_stats)

    # Simulate market matrix (correlation matrix)
    market_matrix = df.corr()
    market_interpretation = generate_market_matrix_interpretation(market_matrix)

    # Perform basic time series analysis (e.g., growth trends)
    growth_trends = df.pct_change().mean()
    growth_interpretation = generate_growth_trend_interpretation(growth_trends)

    # Create visualizations
    pie_chart_path = create_pie_chart(df)
    bar_chart_path = create_bar_chart(df)
    trend_chart_path = create_trend_analysis_chart(df)

    # Create a report file summarizing the analysis
    report_file = os.path.join(app.config['REPORT_FOLDER'], 'report.txt')
    with open(report_file, 'w') as report:
        report.write("InsightGenie Analysis Report\n")
        report.write("\nDescriptive Statistics:\n")
        report.write(descriptive_stats.to_string())
        report.write("\nInterpretation of Descriptive Statistics:\n")
        report.write(descriptive_interpretation)
        report.write("\n\nMarket Matrix (Correlation):\n")
        report.write(market_matrix.to_string())
        report.write("\nInterpretation of Market Matrix:\n")
        report.write(market_interpretation)
        report.write("\n\nGrowth Trends:\n")
        report.write(growth_trends.to_string())
        report.write("\nInterpretation of Growth Trends:\n")
        report.write(growth_interpretation)

    return render_template('report_preview.html',
                           report_filename='report.txt',
                           trend_chart='trend_analysis.png',
                           pie_chart='pie_chart.png',
                           bar_chart='bar_chart.png')

# Route to download report
@app.route('/download/<filename>')
def download_report(filename):
    file_path = os.path.join(app.config['REPORT_FOLDER'], filename)
    return send_file(file_path, as_attachment=True)

# Routes to download specific charts
@app.route('/download/pie_chart')
def download_pie_chart():
    return send_file(os.path.join(app.config['REPORT_FOLDER'], 'pie_chart.png'), as_attachment=True)

@app.route('/download/bar_chart')
def download_bar_chart():
    return send_file(os.path.join(app.config['REPORT_FOLDER'], 'bar_chart.png'), as_attachment=True)

@app.route('/download/trend_chart')
def download_trend_chart():
    return send_file(os.path.join(app.config['REPORT_FOLDER'], 'trend_analysis.png'), as_attachment=True)

def create_pie_chart(df):
    pie_chart_path = os.path.join(app.config['REPORT_FOLDER'], 'pie_chart.png')
    plt.figure(figsize=(8, 8))
    df.iloc[:, 0].value_counts().plot.pie(autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
    plt.title('Distribution of First Column Values', fontsize=16)
    plt.ylabel('')  # Hide the y-label for cleaner look
    plt.grid(axis='y', linestyle='--', alpha=0.7)  # Add gridlines for better readability
    plt.savefig(pie_chart_path, bbox_inches='tight')
    plt.close()
    return pie_chart_path

def create_bar_chart(df):
    bar_chart_path = os.path.join(app.config['REPORT_FOLDER'], 'bar_chart.png')
    plt.figure(figsize=(10, 6))
    df.sum().plot(kind='bar', color='skyblue', edgecolor='black')
    plt.title('Sum of Each Column', fontsize=16)
    plt.ylabel('Sum of Values', fontsize=14)
    plt.xticks(rotation=45)
    plt.grid(axis='y', linestyle='--', alpha=0.7)

    # Reverse the y-axis to have the largest bars at the top
    plt.gca().invert_yaxis()

    plt.savefig(bar_chart_path, bbox_inches='tight')
    plt.close()
    return bar_chart_path

def create_trend_analysis_chart(df):
    trend_chart_path = os.path.join(app.config['REPORT_FOLDER'], 'trend_analysis.png')
    plt.figure(figsize=(12, 6))
    df.plot(title='Trend Analysis', linewidth=2, marker='o')
    plt.title('Trend Analysis Over Time', fontsize=16)
    plt.ylabel('Values', fontsize=14)
    plt.xlabel('Index', fontsize=14)
    plt.grid(axis='both', linestyle='--', alpha=0.7)
    plt.savefig(trend_chart_path, bbox_inches='tight')
    plt.close()
    return trend_chart_path

def generate_interpretation(stats):
    interpretation = ""
    for column in stats.columns:
        interpretation += f"For column '{column}':\n"
        interpretation += f"- Mean: {stats[column]['mean']:.2f}\n"
        interpretation += f"- Standard Deviation: {stats[column]['std']:.2f}\n"
        interpretation += f"- Minimum: {stats[column]['min']}\n"
        interpretation += f"- Maximum: {stats[column]['max']}\n\n"
    return interpretation

def generate_market_matrix_interpretation(matrix):
    interpretation = "Correlation Matrix Interpretation:\n"
    for col in matrix.columns:
        for idx in matrix.index:
            if abs(matrix[col][idx]) > 0.5:
                interpretation += f"There is a significant correlation of {matrix[col][idx]:.2f} between '{col}' and '{idx}'.\n"
    return interpretation

def generate_growth_trend_interpretation(growth_trends):
    interpretation = "Growth Trends Interpretation:\n"
    for col, trend in growth_trends.items():
        interpretation += f"The average growth rate for '{col}' is {trend:.2%}.\n"
    return interpretation

if __name__ == '__main__':
    app.run(debug=True)

