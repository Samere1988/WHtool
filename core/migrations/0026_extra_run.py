from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        (
            "core",
            "0025_billofladingline_package_type",
        ),
    ]

    operations = [
        migrations.CreateModel(
            name="ExtraRun",
            fields=[
                (
                    "id",
                    models.BigAutoField(
                        auto_created=True,
                        primary_key=True,
                        serialize=False,
                        verbose_name="ID",
                    ),
                ),
                (
                    "shipping_date",
                    models.DateField(),
                ),
                (
                    "name",
                    models.CharField(
                        max_length=100,
                    ),
                ),
                (
                    "created_at",
                    models.DateTimeField(
                        auto_now_add=True,
                    ),
                ),
            ],
            options={
                "ordering": [
                    "shipping_date",
                    "created_at",
                    "id",
                ],
            },
        ),

        migrations.AddConstraint(
            model_name="extrarun",
            constraint=models.UniqueConstraint(
                fields=(
                    "shipping_date",
                    "name",
                ),
                name=(
                    "unique_extra_run_name_per_date"
                ),
            ),
        ),
    ]