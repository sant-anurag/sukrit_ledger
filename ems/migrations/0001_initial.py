# Generated by Django 5.0.6 on 2024-05-17 07:14

import django.db.models.deletion
import django.utils.timezone
from django.conf import settings
from django.db import migrations, models


class Migration(migrations.Migration):

    initial = True

    dependencies = [
        ("auth", "0012_alter_user_first_name_max_length"),
    ]

    operations = [
        migrations.CreateModel(
            name="CustomUser",
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
                ("username", models.CharField(max_length=30, unique=True)),
                ("aadhar_number", models.BigIntegerField(blank=True, null=True)),
                (
                    "employment_type",
                    models.CharField(blank=True, max_length=40, null=True),
                ),
                ("password", models.CharField(max_length=100)),
                (
                    "email",
                    models.EmailField(
                        blank=True, max_length=254, null=True, unique=True
                    ),
                ),
                (
                    "mobilenumber",
                    models.CharField(blank=True, max_length=12, null=True, unique=True),
                ),
                ("title", models.CharField(blank=True, max_length=3, null=True)),
                (
                    "firstname",
                    models.CharField(
                        blank=True, max_length=50, null=True, verbose_name="FIrst Name"
                    ),
                ),
                (
                    "lastname",
                    models.CharField(
                        blank=True, max_length=50, null=True, verbose_name="Last Name"
                    ),
                ),
                (
                    "middlename",
                    models.CharField(
                        blank=True,
                        max_length=100,
                        null=True,
                        verbose_name="Middle Name",
                    ),
                ),
                ("first_login", models.BooleanField(default=True)),
                ("last_login", models.DateTimeField(blank=True, null=True)),
                (
                    "password_changed_by",
                    models.CharField(
                        blank=True,
                        max_length=20,
                        null=True,
                        verbose_name="Password Changed By",
                    ),
                ),
                ("password_changed", models.BooleanField(default=False)),
                ("password_changed_date", models.DateTimeField(blank=True, null=True)),
                ("is_superuser", models.BooleanField(default=False)),
                (
                    "last_password_reset_date",
                    models.DateField(
                        auto_now_add=True,
                        null=True,
                        verbose_name="Last Password Reset Date",
                    ),
                ),
                (
                    "next_password_reset_date",
                    models.DateField(
                        auto_now_add=True,
                        null=True,
                        verbose_name="Next Password Reset Date",
                    ),
                ),
                ("is_staff", models.BooleanField(default=False)),
                ("is_active", models.BooleanField(default=True)),
                (
                    "is_deleted",
                    models.BooleanField(default=False, verbose_name="Is Deleted"),
                ),
                (
                    "backup_deleted_by",
                    models.CharField(
                        max_length=55, null=True, verbose_name="Backup Deleted By"
                    ),
                ),
                ("deleted_at", models.DateTimeField(null=True)),
                (
                    "created_at",
                    models.DateTimeField(
                        default=django.utils.timezone.now,
                        editable=False,
                        verbose_name="Created At",
                    ),
                ),
                (
                    "modified_at",
                    models.DateTimeField(null=True, verbose_name="Modified At"),
                ),
                (
                    "created_by",
                    models.ForeignKey(
                        null=True,
                        on_delete=django.db.models.deletion.SET_NULL,
                        related_name="customuser_creater",
                        to=settings.AUTH_USER_MODEL,
                        verbose_name="Created By",
                    ),
                ),
                (
                    "deleted_by",
                    models.ForeignKey(
                        null=True,
                        on_delete=django.db.models.deletion.SET_NULL,
                        related_name="ems_deletor",
                        to=settings.AUTH_USER_MODEL,
                    ),
                ),
                (
                    "groups",
                    models.ManyToManyField(
                        blank=True,
                        help_text="The groups this user belongs to. A user will get all permissions granted to each of their groups.",
                        related_name="user_set",
                        related_query_name="user",
                        to="auth.group",
                        verbose_name="groups",
                    ),
                ),
                (
                    "modified_by",
                    models.ForeignKey(
                        null=True,
                        on_delete=django.db.models.deletion.SET_NULL,
                        related_name="customuser_modifier",
                        to=settings.AUTH_USER_MODEL,
                        verbose_name="Modified By",
                    ),
                ),
                (
                    "user_permissions",
                    models.ManyToManyField(
                        blank=True,
                        help_text="Specific permissions for this user.",
                        related_name="user_set",
                        related_query_name="user",
                        to="auth.permission",
                        verbose_name="user permissions",
                    ),
                ),
            ],
            options={
                "abstract": False,
            },
        ),
    ]