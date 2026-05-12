# Generated manually for editable transport driver start times

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0010_transport_import_workflow'),
    ]

    operations = [
        migrations.AddField(
            model_name='runsheet',
            name='transport_start_time',
            field=models.CharField(blank=True, max_length=50, null=True),
        ),
        migrations.AddField(
            model_name='transportimportrow',
            name='imported_start_time',
            field=models.CharField(blank=True, max_length=50),
        ),
        migrations.AddField(
            model_name='transportimportpreviousstate',
            name='previous_transport_start_time',
            field=models.CharField(blank=True, max_length=50, null=True),
        ),
    ]
