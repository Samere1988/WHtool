# Generated manually for transport-company import review/apply/undo workflow

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0009_alter_finalizedrunsheet_table'),
    ]

    operations = [
        migrations.CreateModel(
            name='TransportImportBatch',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('shipping_date', models.DateField()),
                ('original_filename', models.CharField(blank=True, max_length=255)),
                ('uploaded_by', models.CharField(blank=True, max_length=150)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('applied_at', models.DateTimeField(blank=True, null=True)),
                ('undone_at', models.DateTimeField(blank=True, null=True)),
                ('status', models.CharField(choices=[('review', 'Review'), ('applied', 'Applied'), ('undone', 'Undone'), ('failed', 'Failed')], default='review', max_length=20)),
                ('notes', models.TextField(blank=True)),
            ],
            options={
                'ordering': ['-created_at'],
            },
        ),
        migrations.CreateModel(
            name='TransportImportPreviousState',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('run_sheet_id', models.IntegerField()),
                ('previous_transport_run_name', models.CharField(blank=True, max_length=100, null=True)),
                ('previous_transport_driver', models.CharField(blank=True, max_length=100, null=True)),
                ('previous_transport_truck', models.CharField(blank=True, max_length=100, null=True)),
                ('previous_transport_stop_number', models.IntegerField(blank=True, null=True)),
                ('previous_transport_import_batch_id', models.IntegerField(blank=True, null=True)),
                ('previous_driver_name', models.CharField(blank=True, max_length=100, null=True)),
                ('previous_load_index', models.IntegerField(default=0)),
                ('batch', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='previous_states', to='core.transportimportbatch')),
            ],
            options={
                'unique_together': {('batch', 'run_sheet_id')},
            },
        ),
        migrations.CreateModel(
            name='TransportImportRow',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('sort_order', models.IntegerField(default=0)),
                ('source_sheet_name', models.CharField(blank=True, max_length=100)),
                ('source_row_number', models.IntegerField(default=0)),
                ('imported_run_name', models.CharField(blank=True, max_length=100)),
                ('imported_driver', models.CharField(blank=True, max_length=100)),
                ('imported_truck', models.CharField(blank=True, max_length=100)),
                ('imported_stop_number', models.IntegerField(default=0)),
                ('imported_customer_name', models.CharField(blank=True, max_length=255)),
                ('imported_city', models.CharField(blank=True, max_length=150)),
                ('matched_run_sheet_ids', models.TextField(blank=True)),
                ('confidence', models.FloatField(default=0)),
                ('status', models.CharField(choices=[('matched', 'Matched'), ('review', 'Needs Review'), ('unmatched', 'Unmatched')], default='unmatched', max_length=20)),
                ('batch', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='rows', to='core.transportimportbatch')),
            ],
            options={
                'ordering': ['sort_order', 'id'],
            },
        ),
        migrations.AddField(
            model_name='runsheet',
            name='transport_driver',
            field=models.CharField(blank=True, max_length=100, null=True),
        ),
        migrations.AddField(
            model_name='runsheet',
            name='transport_import_batch',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='run_sheet_items', to='core.transportimportbatch'),
        ),
        migrations.AddField(
            model_name='runsheet',
            name='transport_run_name',
            field=models.CharField(blank=True, max_length=100, null=True),
        ),
        migrations.AddField(
            model_name='runsheet',
            name='transport_stop_number',
            field=models.IntegerField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name='runsheet',
            name='transport_truck',
            field=models.CharField(blank=True, max_length=100, null=True),
        ),
    ]
