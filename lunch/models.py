
from django.db import models
from django.contrib.auth import get_user_model

User = get_user_model()
class LunchConfig(models.Model):
    price         = models.IntegerField(default=650, help_text="お弁当価格")
    subsidy       = models.IntegerField(default=230, help_text="会社補助額")
    monthly_limit = models.IntegerField(default=3780, help_text="会社負担上限額")

    class Meta:
        verbose_name        = "ランチ設定"
        verbose_name_plural = "ランチ設定"

    def __str__(self):
        return f"ランチ設定 (価格:{self.price} 補助:{self.subsidy} 上限:{self.monthly_limit})"

class Order(models.Model):
    RICE_SIZES = [
        ('大', 'ライス大'),
        ('中', 'ライス中'),
        ('小', 'ライス小'),
    ]
    VENDORS = [
        ('veg17',   'ベジタブルディッシュ17'),
        ('yamajin', 'やまじん'),
        ('kaachan','かあちゃんの台所'),
    ]

    user        = models.ForeignKey(User, on_delete=models.CASCADE)
    order_date  = models.DateField()
    vendor      = models.CharField(max_length=20, choices=VENDORS)
    rice_size   = models.CharField(max_length=2, choices=RICE_SIZES, default='中')
    quantity    = models.PositiveIntegerField(default=1)
    price       = models.IntegerField(default=650)
    subsidy     = models.IntegerField(default=230)
    status      = models.CharField(
        max_length=20,
        choices=[
            ('pending', '未発注'),
            ('sent',    '発注済'),
        ],
        default='pending',
        verbose_name="ステータス",
    )
    canceled    = models.BooleanField(default=False)
    canceled_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        unique_together = ('user','order_date','vendor','rice_size')
