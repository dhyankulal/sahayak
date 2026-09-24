from django.db import migrations, models

class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='PoliceReport',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('case_number', models.CharField(max_length=100)),
                ('type_of_case', models.CharField(max_length=150)),
                ('age_of_person', models.IntegerField(blank=True, null=True)),
                ('police_name', models.CharField(max_length=150)),
                ('address', models.CharField(max_length=300)),
                ('volunteer_name', models.CharField(blank=True, max_length=150, null=True)),
                ('phone_number', models.CharField(max_length=50)),
                ('type_of_help_needed', models.TextField()),
                ('actions_taken', models.TextField()),
                ('date_of_report', models.DateField(auto_now_add=True)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
            ],
        ),
    ]
