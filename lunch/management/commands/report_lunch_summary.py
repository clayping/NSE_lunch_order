import calendar
from datetime import date
from django.core.management.base import BaseCommand
from django.contrib.auth import get_user_model
from lunch.models import Order, LunchConfig
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

class Command(BaseCommand):
    help = "月末ランチ注文レポートを Excel で出力します"

    def add_arguments(self, parser):
        parser.add_argument('--year',  type=int, default=date.today().year)
        parser.add_argument('--month', type=int, default=date.today().month)

    def handle(self, *args, **options):
        # ── LunchConfig の取得 or 作成 ──
        cfg, created = LunchConfig.objects.get_or_create(
            defaults={'price': 650, 'subsidy': 230, 'monthly_limit': 3780}
        )
        if created:
            self.stdout.write(self.style.WARNING(
                'LunchConfig レコードが存在しなかったため、デフォルト値で自動作成しました'
            ))

        year  = options['year']
        month = options['month']
        User  = get_user_model()
        days  = calendar.monthrange(year, month)[1]
        users = User.objects.all().order_by('username')

        # ワークブック＆シート
        wb  = Workbook()
        ws  = wb.active
        ws.title = f"{year}年{month}月ランチ注文"

        # 列ヘッダー
        header = ['コード','氏名'] \
               + [f"{d}日" for d in range(1, days+1)] \
               + ['注文数','合計金額','補助額','上限','会社負担','超過','実費']
        ws.append(header)

        # ヘッダー書式
        for col in range(1, len(header)+1):
            cell = ws.cell(row=1, column=col)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="DDDDDD")
            cell.alignment = Alignment(horizontal='center')

        # 各ユーザー行
        for row_idx, user in enumerate(users, start=2):
            code = user.id
            name = user.get_full_name() or user.username

            # 日別フラグ
            flags = [
                1 if Order.objects.filter(
                        user=user,
                        order_date=date(year, month, d),
                        canceled=False
                   ).exists() else 0
                for d in range(1, days+1)
            ]

            # 集計
            total_qty     = sum(flags)
            total_price   = total_qty * cfg.price
            total_subsidy = total_qty * cfg.subsidy
            limit         = cfg.monthly_limit
            company_pay   = min(total_subsidy, limit)
            over          = max(0, total_subsidy - limit)
            user_pay      = total_price - company_pay

            row = [code, name] + flags + [
                total_qty,
                total_price,
                total_subsidy,
                limit,
                company_pay,
                over,
                user_pay,
            ]
            ws.append(row)

        # 日別合計行
        total_row = ['', '合計']
        # 日別合計
        for d in range(1, days+1):
            cnt = Order.objects.filter(
                order_date=date(year, month, d),
                canceled=False
            ).count()
            total_row.append(cnt)
        # 合計列集計
        total_qty     = sum(total_row[2:2+days])
        total_price   = total_qty * cfg.price
        total_subsidy = total_qty * cfg.subsidy
        limit         = cfg.monthly_limit
        company_pay   = min(total_subsidy, limit)
        over          = max(0, total_subsidy - limit)
        user_pay      = total_price - company_pay
        total_row += [
            total_qty,
            total_price,
            total_subsidy,
            limit,
            company_pay,
            over,
            user_pay,
        ]
        ws.append(total_row)

        # 合計行書式
        last_row = ws.max_row
        for col in range(1, len(header)+1):
            cell = ws.cell(row=last_row, column=col)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="FFFF99")

        # 下部にベンダー集計
        start = ws.max_row + 2
        ws.cell(row=start, column=1, value='≪ベンダー集計≫').font = Font(bold=True)
        for i, (code, name) in enumerate(Order.VENDORS, start= start+1):
            qs = Order.objects.filter(
                order_date__year=year,
                order_date__month=month,
                vendor=code,
                canceled=False
            )
            cnt = qs.count()
            amt = sum(o.price for o in qs)
            ws.cell(row=i, column=2, value=name)
            ws.cell(row=i, column=days+3, value=cnt)
            ws.cell(row=i, column=days+4, value=amt)

        # 保存
        filename = f'lunch_report_{year}{month:02}.xlsx'
        wb.save(filename)
        self.stdout.write(self.style.SUCCESS(f'レポートを {filename} に出力しました'))
