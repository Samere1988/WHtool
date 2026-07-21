from django.db.models.signals import post_delete
from django.dispatch import receiver

from .models import ContainerPhoto, OutboundPhoto, PickupPhoto


@receiver(
    post_delete,
    sender=ContainerPhoto,
)
@receiver(
    post_delete,
    sender=OutboundPhoto,
)
@receiver(
    post_delete,
    sender=PickupPhoto,
)
def delete_photo_file(sender, instance, **kwargs):
    image = getattr(instance, "image", None)

    if image and image.name:
        image.delete(save=False)
