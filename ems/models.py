from django.db import models
from django.contrib.auth.models import AbstractBaseUser, PermissionsMixin
from django.db import models
from django.utils import timezone
from .managers import CustomUserManager

# Create your models here.

class CustomUser(AbstractBaseUser, PermissionsMixin):
    username = models.CharField(unique=True, max_length=30)
    aadhar_number = models.BigIntegerField(null=True, blank=True)
    employment_type = models.CharField(max_length=40, blank=True, null=True)
    password = models.CharField(max_length=100, null=False)
    email = models.EmailField(null=True, blank=True, unique=True)
    mobilenumber = models.CharField(max_length=12, unique=True, blank=True, null=True)
    
    title = models.CharField(max_length=3,blank=True,null=True)
    firstname = models.CharField(max_length=50, verbose_name="FIrst Name", blank=True, null=True)
    lastname = models.CharField(max_length=50, verbose_name="Last Name", blank=True, null=True)
    middlename = models.CharField(max_length=100,null=True,blank=True,verbose_name='Middle Name')
    
    first_login = models.BooleanField(default=True)
    last_login = models.DateTimeField(blank=True, null=True)
    password_changed_by = models.CharField(max_length=20,null=True,blank=True,verbose_name='Password Changed By')
    password_changed = models.BooleanField(default=False)
    password_changed_date = models.DateTimeField(blank=True, null=True)
    is_superuser = models.BooleanField(default=False)
    last_password_reset_date     = models.DateField(auto_now_add=True,blank=True,null=True,verbose_name='Last Password Reset Date')
    next_password_reset_date     = models.DateField(auto_now_add=True,blank=True,null=True,verbose_name='Next Password Reset Date')
    is_staff = models.BooleanField(default=False)
    is_active = models.BooleanField(default=True)
    is_deleted   = models.BooleanField(default=False,verbose_name="Is Deleted")
    deleted_by   = models.ForeignKey("self", on_delete=models.SET_NULL, null=True, related_name="ems_deletor")
    backup_deleted_by = models.CharField(max_length=55, null=True, verbose_name="Backup Deleted By")
    deleted_at   = models.DateTimeField(null=True)
    created_at   = models.DateTimeField(editable=False,default=timezone.now,verbose_name='Created At')
    modified_at  = models.DateTimeField(null=True,verbose_name='Modified At')
    created_by   = models.ForeignKey('self', on_delete=models.SET_NULL, null=True,verbose_name='Created By', related_name="customuser_creater")
    modified_by  = models.ForeignKey('self', on_delete=models.SET_NULL, null=True,verbose_name='Modified By', related_name="customuser_modifier")


    USERNAME_FIELD = 'username'
    REQUIRED_FIELDS = []

    objects = CustomUserManager()

    def __str__(self):
        return self.username
