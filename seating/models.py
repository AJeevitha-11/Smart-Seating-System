from django.db import models

class Block(models.Model):
    name = models.CharField(max_length=10, unique=True)

    def __str__(self):
        return self.name


class Room(models.Model):
    block = models.ForeignKey(Block, on_delete=models.CASCADE)
    room_number = models.CharField(max_length=20)
    rows = models.IntegerField()
    columns = models.IntegerField()
    capacity = models.IntegerField()

    def save(self, *args, **kwargs):
        # auto-calculate capacity if not set
        if not self.capacity:
            self.capacity = self.rows * self.columns
        super().save(*args, **kwargs)

    def __str__(self):
        return f"{self.block.name} - {self.room_number}"



class ExamSlot(models.Model):
    name = models.CharField(max_length=100)
    date = models.DateField()
    time_range = models.CharField(max_length=50)  # e.g., 10AM - 12PM

    def __str__(self):
        return f"{self.name} - {self.date} {self.time_range}"


