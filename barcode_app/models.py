from django.db import models


class HAWBRecord(models.Model):
    """Stores every HAWB record uploaded via XML."""

    mawb          = models.CharField(max_length=50,  blank=True, default='')
    run_no        = models.CharField(max_length=50,  blank=True, default='')
    hawb_number   = models.CharField(max_length=20,  unique=True)
    no_of_bags    = models.CharField(max_length=20,  blank=True, default='')
    weight        = models.CharField(max_length=20,  blank=True, default='')
    destination   = models.CharField(max_length=10,  blank=True, default='')
    iec           = models.CharField(max_length=30,  blank=True, default='')
    mhbs_no       = models.CharField(max_length=50,  blank=True, default='')
    consignor     = models.CharField(max_length=200, blank=True, default='')
    consignee     = models.CharField(max_length=200, blank=True, default='')
    invoice_no    = models.CharField(max_length=50,  blank=True, default='')
    invoice_date  = models.CharField(max_length=30,  blank=True, default='')
    gstin         = models.CharField(max_length=30,  blank=True, default='')
    flight_no     = models.CharField(max_length=20,  blank=True, default='')
    departure_date = models.CharField(max_length=30, blank=True, default='')
    status        = models.CharField(max_length=50,  blank=True, default='')
    is_special    = models.BooleanField(default=False)

    class Meta:
        ordering = ['is_special', 'hawb_number']
        verbose_name = 'HAWB Record'
        verbose_name_plural = 'HAWB Records'

    def __str__(self):
        tag = ' ★' if self.is_special else ''
        return f"{self.hawb_number}{tag} | {self.run_no} | {self.mawb}"


class HAWBState(models.Model):
    """Singleton row tracking the last generated 9-digit HAWB serial number."""

    last_generated = models.BigIntegerField(default=0)
    last_batch_start = models.BigIntegerField(default=0)
    last_batch_count = models.IntegerField(default=0)

    class Meta:
        verbose_name = 'HAWB State'

    def __str__(self):
        return f"Last Generated: {self.last_generated}"

    @classmethod
    def get(cls):
        obj, _ = cls.objects.get_or_create(pk=1)
        return obj
