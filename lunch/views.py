# from datetime import date
import calendar
import json
from datetime import date, time, timedelta
from django.shortcuts import render, redirect
from django.http import HttpResponse
from django.template.loader import render_to_string
from django.utils import timezone
from django.contrib.auth.decorators import login_required
from weasyprint import HTML
from django.http import JsonResponse
from django.views.decorators.http import require_POST

from .models import Order


def get_allowed_dates(start: date, count: int) -> set[date]:
    """
    start日から「日曜を除く」count 日間を集めてセットで返す。
    """
    allowed = set()
    d = start
    # count 分だけ日曜以外をcollect
    while len(allowed) < count:
        if d.weekday() != 6:  # 6=日曜
            allowed.add(d)
        d += timedelta(days=1)
    return allowed

@login_required
def monthly_calendar(request, year=None, month=None):
    """月間カレンダー表示ビュー"""
    today = date.today()
    year  = year or today.year
    month = month or today.month

    # ── 追加: 許可日を計算 ──
    allowed_dates = get_allowed_dates(today, 6)
    # テンプレートでも today と allowed_dates を参照できるように渡す

    cal = calendar.Calendar(firstweekday=0)
    month_days = cal.monthdatescalendar(year, month)

    # ユーザーの今月の注文日をセット化
    orders = Order.objects.filter(
        user=request.user,
        order_date__year=year,
        order_date__month=month,
        canceled=False
    ).values_list('order_date', flat=True)
    orders_set = set(orders)

    # 日付ごとに「注文済みか」を付与
    calendar_data = []
    for week in month_days:
        week_data = []
        for day in week:
            week_data.append({
                'day': day,
                'is_current_month': day.month == month,
                'ordered': day in orders_set,
                'allowed': day in allowed_dates,
            })
        calendar_data.append(week_data)

    return render(request, 'lunch/calendar.html', {
        'calendar_data': calendar_data,
        'year': year,
        'month': month,
        'today': today,
        'allowed_dates': allowed_dates,
    })
def fax_order_pdf(request):
    # 今日は何件注文があるか集計
    today = date.today()
    qs = Order.objects.filter(order_date=today, canceled=False)
    # ライスの大中小それぞれ注文数を仮に集計する例
    counts = {
        'large':  qs.filter(rice_size='大').count(),
        'medium': qs.filter(rice_size='中').count(),
        'small':  qs.filter(rice_size='小').count(),
    }

    # HTML をレンダリングして PDF 化
    html_string = render_to_string('lunch/fax_template.html', {
        'today': today,
        'counts': counts,
    })
    html = HTML(string=html_string)
    pdf = html.write_pdf()

    # レスポンスで返却
    response = HttpResponse(pdf, content_type='application/pdf')
    filename = today.strftime('lunch_order_%Y%m%d.pdf')
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    return response

@login_required
def today_order(request):
    """当日の注文を行う／キャンセルする画面"""
    user = request.user
    today = date.today()

    # まず、当日の注文レコードをキャンセルフラグに関係なく取得
    order = Order.objects.filter(user=user, order_date=today).first()
    
    if request.method == 'POST':
        action = request.POST.get('action')
        if action == 'order':
            if not order:
                # レコード自体がなければ新規作成
                Order.objects.create(user=user, order_date=today, vendor='veg17', rice_size='中')
            else:
                # 既存レコードがあるならキャンセルフラグをオフに
                order.canceled = False
                order.canceled_at = None
                order.save()
        elif action == 'cancel' and order:
            # キャンセルするときはフラグをオンにして日時を埋める
            order.canceled = True
            order.canceled_at = timezone.now()
            order.save()
        return redirect('today_order')

    # ★キャンセル済みの場合はテンプレート上では「注文なし扱い」にする
    order_for_template = order if order and not order.canceled else None
    return render(request, 'lunch/today_order.html', {
        'order': order_for_template,
        'today': today,
    })
@login_required
@require_POST
def toggle_order(request):
    """
    POST JSON { "date": "YYYY-MM-DD" }
    → その日の注文レコードを必ず取得 or 作成し、canceled フラグをトグル
    """
    data = json.loads(request.body)
    try:
        day = date.fromisoformat(data['date'])
    except Exception:
        return JsonResponse({'error': 'invalid date'}, status=400)

    # 当日9:00以降は締め切り
    now = timezone.localtime()
    # 当日から日曜除く6日以外は変更不可
    allowed = get_allowed_dates(date.today(), 6)
    if day not in allowed:
        return JsonResponse({'error': '変更可能期間外です'}, status=403)

    if now.time() >= time(8, 10) and day == date.today():
        return JsonResponse({'error': '受付は午前9時までです'}, status=403)
    # get_or_create ならレコードがなければ作ってくれる
    order, created = Order.objects.get_or_create(
        user=request.user,
        order_date=day,
        defaults={'vendor': 'veg17', 'rice_size': '中'}
    )

    # 事務がステータスを発注済にした後は変更不可
    if hasattr(order, 'status') and order.status != 'pending':
        return JsonResponse({'error': '既に発注済のため変更できません'}, status=403)
    # トグル処理
    if order.canceled:
        order.canceled = False
        order.canceled_at = None
        status = 'ordered'
    else:
        order.canceled = True
        order.canceled_at = timezone.now()
        status = 'canceled'

    order.save()
    return JsonResponse({'status': status, 'date': data['date']})
