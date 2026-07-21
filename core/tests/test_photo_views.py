import shutil
import tempfile
from datetime import date
from io import BytesIO
from pathlib import Path

from django.contrib.auth import get_user_model
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase, override_settings
from django.urls import reverse
from PIL import Image

from core.models import Container, ContainerPhoto


TEST_MEDIA_ROOT = tempfile.mkdtemp(
    prefix="whtool-photo-tests-",
)


def make_test_photo(name="test-photo.jpg"):
    output = BytesIO()

    Image.new(
        "RGB",
        (40, 30),
        "navy",
    ).save(
        output,
        format="JPEG",
    )

    return SimpleUploadedFile(
        name,
        output.getvalue(),
        content_type="image/jpeg",
    )


@override_settings(MEDIA_ROOT=TEST_MEDIA_ROOT)
class ContainerPhotoViewTests(TestCase):
    @classmethod
    def tearDownClass(cls):
        super().tearDownClass()
        shutil.rmtree(
            TEST_MEDIA_ROOT,
            ignore_errors=True,
        )

    def setUp(self):
        user = get_user_model().objects.create_user(
            username="photo-test-user",
            password="test-password",
        )

        self.client.force_login(user)

    def test_add_container_keeps_selected_date(self):
        response = self.client.post(
            reverse("add_container"),
            {
                "container_number": "TEST-123",
                "date_received": "2026-07-01",
                "unloaded_by": "Tester",
                "unloaded_at": "20 Hymus",
                "photos": make_test_photo(),
            },
        )

        container = Container.objects.get(
            container_number="TEST-123",
        )

        self.assertRedirects(
            response,
            reverse(
                "container_detail",
                kwargs={"pk": container.pk},
            ),
        )

        self.assertEqual(
            container.date_received,
            date(2026, 7, 1),
        )

        self.assertEqual(
            container.photos.count(),
            1,
        )

    def test_delete_photo_requires_post_and_deletes_file(self):
        container = Container.objects.create(
            container_number="TEST-DELETE",
            date_received=date(2026, 7, 1),
            unloaded_by="Tester",
            unloaded_at="20 Hymus",
        )

        photo = ContainerPhoto.objects.create(
            container=container,
            image=make_test_photo("delete-me.jpg"),
        )

        photo_path = Path(photo.image.path)
        delete_url = reverse(
            "delete_container_photo",
            kwargs={"photo_id": photo.pk},
        )

        self.assertTrue(photo_path.exists())
        self.assertEqual(
            self.client.get(delete_url).status_code,
            405,
        )

        response = self.client.post(delete_url)

        self.assertRedirects(
            response,
            reverse(
                "container_detail",
                kwargs={"pk": container.pk},
            ),
        )

        self.assertFalse(
            ContainerPhoto.objects.filter(
                pk=photo.pk,
            ).exists()
        )

        self.assertFalse(photo_path.exists())

    def test_detail_page_has_previous_and_next_controls(self):
        container = Container.objects.create(
            container_number="TEST-GALLERY",
            date_received=date(2026, 7, 1),
            unloaded_by="Tester",
            unloaded_at="20 Hymus",
        )

        ContainerPhoto.objects.create(
            container=container,
            image=make_test_photo("one.jpg"),
        )

        ContainerPhoto.objects.create(
            container=container,
            image=make_test_photo("two.jpg"),
        )

        response = self.client.get(
            reverse(
                "container_detail",
                kwargs={"pk": container.pk},
            )
        )

        self.assertContains(
            response,
            'id="previousPhotoButton"',
        )

        self.assertContains(
            response,
            'id="nextPhotoButton"',
        )

        self.assertContains(
            response,
            'data-photo-index="0"',
        )

        self.assertContains(
            response,
            'data-photo-index="1"',
        )
