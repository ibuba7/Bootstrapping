import io
import os
import pandas as pd
import base64
from django.shortcuts import render
from django.http import JsonResponse
import matplotlib.pyplot as plt
from .curve_bootstrapping import YieldCurve
import matplotlib.ticker as mtick

def home(request):
    return render(request, 'calculator/home.html')

def bootstrapping(request):
    irs = YieldCurve()
    irs.BootstrapYieldCurve()

    # Generate Zero Curve plot
    fig, ax = plt.subplots()
    ax.plot(irs.dfcurve.Date, irs.dfcurve.ZeroRate, marker='o')
    ax.grid()
    ax.set_title('Zero Curve')
    ax.yaxis.set_major_formatter(mtick.PercentFormatter(1.0))  # Format y-axis as percentage
    plt.xticks(rotation=90)
    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    buf.seek(0)
    zero_curve_img = base64.b64encode(buf.read()).decode('utf-8')
    plt.close(fig)

    # Generate Discount Factor plot
    fig, ax = plt.subplots()
    ax.plot(irs.dfcurve.Date, irs.dfcurve.DiscountFactor, marker='o')
    ax.grid()
    ax.set_title('Discount Factor')
    ax.yaxis.set_major_formatter(mtick.PercentFormatter(1.0))  # Format y-axis as percentage
    plt.xticks(rotation=90)
    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    buf.seek(0)
    discount_factor_img = base64.b64encode(buf.read()).decode('utf-8')
    plt.close(fig)

    # Generate Forward Rate plot
    fig, ax = plt.subplots()
    ax.plot(irs.dfcurve.Date, irs.dfcurve.ForwardRate, marker='o')
    ax.grid()
    ax.set_title('Forward Rate')
    ax.yaxis.set_major_formatter(mtick.PercentFormatter(1.0))  # Format y-axis as percentage
    plt.xticks(rotation=90)
    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    buf.seek(0)
    forward_rate_img = base64.b64encode(buf.read()).decode('utf-8')
    plt.close(fig)

    # Read data from the "Cetes" and "Mbonos" sheets
    filein = os.path.join("sofr_data.xlsx")
    df_cetes = pd.read_excel(filein, sheet_name='Cetes')
    df_mbonos = pd.read_excel(filein, sheet_name='Mbonos')

    # Format table values as percentages
    df_cetes['Nivel'] = df_cetes['Nivel'].apply(lambda x: f'{x * 100:.2f}%')
    df_mbonos['Actual'] = df_mbonos['Actual'].apply(lambda x: f'{x * 100:.2f}%')

    # Generate Cetes plot
    fig, ax = plt.subplots()
    ax.plot(df_cetes['Plazo (Días)'], df_cetes['Nivel'], marker='o')
    ax.grid()
    ax.set_title('Cetes Curve')
    ax.yaxis.set_major_formatter(mtick.PercentFormatter(1.0))  # Format y-axis as percentage
    plt.xticks(rotation=90)
    plt.xlabel('Plazo (Días)')
    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    buf.seek(0)
    cetes_curve_img = base64.b64encode(buf.read()).decode('utf-8')
    plt.close(fig)

    # Generate Mbonos plot
    fig, ax = plt.subplots()
    ax.plot(df_mbonos['Plazo (Días)'], df_mbonos['Actual'], marker='o')
    ax.grid()
    ax.set_title('Mbonos Curve')
    ax.yaxis.set_major_formatter(mtick.PercentFormatter(1.0))  # Format y-axis as percentage
    plt.xticks(rotation=90)
    plt.xlabel('Plazo (Días)')
    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    buf.seek(0)
    mbonos_curve_img = base64.b64encode(buf.read()).decode('utf-8')
    plt.close(fig)

    # Return the images and data as JSON
    return JsonResponse({
        'zero_curve_img': zero_curve_img,
        'discount_factor_img': discount_factor_img,
        'forward_rate_img': forward_rate_img,
        'cetes_curve_img': cetes_curve_img,
        'mbonos_curve_img': mbonos_curve_img,
        'cetes_data': df_cetes.to_html(index=False, classes='table table-striped table-bordered text-center'),
        'mbonos_data': df_mbonos.to_html(index=False, classes='table table-striped table-bordered text-center')
    })