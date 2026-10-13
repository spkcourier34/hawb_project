from django.db import migrations, models


class Migration(migrations.Migration):

    initial = True

    dependencies = []

    operations = [
        migrations.CreateModel(
            name='HAWBRecord',
            fields=[
                ('id',             models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('mawb',           models.CharField(blank=True, default='', max_length=50)),
                ('run_no',         models.CharField(blank=True, default='', max_length=50)),
                ('hawb_number',    models.CharField(max_length=20, unique=True)),
                ('no_of_bags',     models.CharField(blank=True, default='', max_length=20)),
                ('weight',         models.CharField(blank=True, default='', max_length=20)),
                ('destination',    models.CharField(blank=True, default='', max_length=10)),
                ('iec',            models.CharField(blank=True, default='', max_length=30)),
                ('mhbs_no',        models.CharField(blank=True, default='', max_length=50)),
                ('consignor',      models.CharField(blank=True, default='', max_length=200)),
                ('consignee',      models.CharField(blank=True, default='', max_length=200)),
                ('invoice_no',     models.CharField(blank=True, default='', max_length=50)),
                ('invoice_date',   models.CharField(blank=True, default='', max_length=30)),
                ('gstin',          models.CharField(blank=True, default='', max_length=30)),
                ('flight_no',      models.CharField(blank=True, default='', max_length=20)),
                ('departure_date', models.CharField(blank=True, default='', max_length=30)),
                ('is_special',     models.BooleanField(default=False)),
            ],
            options={
                'verbose_name': 'HAWB Record',
                'verbose_name_plural': 'HAWB Records',
                'ordering': ['is_special', 'hawb_number'],
            },
        ),
        migrations.CreateModel(
            name='HAWBState',
            fields=[
                ('id',             models.AutoField(primary_key=True, serialize=False)),
                ('last_generated', models.BigIntegerField(default=0)),
            ],
            options={
                'verbose_name': 'HAWB State',
            },
        ),
    ]
