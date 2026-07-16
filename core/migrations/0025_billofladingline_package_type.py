from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        (
            "core",
            "0024_merge_20260714_1500",
        ),
    ]

    operations = [
        migrations.AddField(
            model_name="billofladingline",
            name="package_type",
            field=models.CharField(
                blank=True,
                choices=[
                    ("skid", "Skid(s)"),
                    ("coil", "Coil(s)"),
                    ("bundle", "Bundle(s)"),
                ],
                max_length=10,
            ),
        ),
    ]