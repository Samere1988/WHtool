import django.utils.timezone
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("core", "0026_extra_run"),
    ]

    operations = [
        migrations.AlterField(
            model_name="container",
            name="date_received",
            field=models.DateField(
                default=django.utils.timezone.localdate,
            ),
        ),
        migrations.AlterField(
            model_name="outboundload",
            name="date_loaded",
            field=models.DateField(
                default=django.utils.timezone.localdate,
            ),
        ),
        migrations.AlterField(
            model_name="pickupphotolog",
            name="date_picked_up",
            field=models.DateField(
                default=django.utils.timezone.localdate,
            ),
        ),
        migrations.AlterField(
            model_name="pickupphotolog",
            name="order_number",
            field=models.CharField(
                blank=True,
                default="",
                max_length=100,
            ),
        ),
    ]
