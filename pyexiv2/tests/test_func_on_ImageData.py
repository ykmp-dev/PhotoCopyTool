from .base import *


def setup_function():
    if ENV.skip_test:
        pytest.skip()
    shutil.copy(ENV.jpg_img, ENV.test_img)  # Before each test, make a temporary copy of the image
    with open(ENV.test_img, 'rb') as f:
        ENV.img = ImageData(f.read())


def test_modify_exif():
    with open(ENV.test_img, 'rb+') as f:
        with ImageData(f.read()) as img:
            changes = {'Exif.Image.ImageDescription': 'test-中文-',
                       'Exif.Image.Artist': None}
            img.modify_exif(changes)
            f.seek(0)
            f.write(img.get_bytes())
        f.seek(0)
        with ImageData(f.read()) as img:
            expected_result = simulate_updating_metadata(data.EXIF, changes)
            result = img.read_exif()
            ignored_keys = ['Exif.Image.ExifTag', 'Exif.Thumbnail.JPEGInterchangeFormat']
            for key in ignored_keys:
                expected_result.pop(key)
                result.pop(key)
            diff_dict(expected_result, result)


def test_modify_iptc():
    with open(ENV.test_img, 'rb+') as f:
        with ImageData(f.read()) as img:
            changes = {'Iptc.Application2.ObjectName': 'test-中文-',
                       'Iptc.Application2.Copyright': None,
                       'Iptc.Application2.Keywords': ['tag1', 'tag2', 'tag3']}
            img.modify_iptc(changes)
            f.seek(0)
            f.write(img.get_bytes())
        f.seek(0)
        with ImageData(f.read()) as img:
            expected_result = simulate_updating_metadata(data.IPTC, changes)
            diff_dict(expected_result, img.read_iptc())


def test_modify_xmp():
    with open(ENV.test_img, 'rb+') as f:
        with ImageData(f.read()) as img:
            changes = {'Xmp.xmp.CreateDate': '2019-06-23T19:45:17.834',
                       'Xmp.xmp.Rating': None,
                       'Xmp.dc.subject': ['tag1', 'tag2', 'tag3']}
            img.modify_xmp(changes)
            f.seek(0)
            f.write(img.get_bytes())
        f.seek(0)
        with ImageData(f.read()) as img:
            expected_result = simulate_updating_metadata(data.XMP, changes)
            diff_dict(expected_result, img.read_xmp())


def test_clear_all():
    with open(ENV.test_img, 'rb+') as f:
        with ImageData(f.read()) as img:
            img.clear_exif()
            img.clear_iptc()
            img.clear_xmp()
            img.clear_comment()
            img.clear_icc()
            img.clear_thumbnail()
            f.seek(0)
            f.write(img.get_bytes())
        f.seek(0)
        with ImageData(f.read()) as img:
            assert img.read_exif() == {}
            assert img.read_iptc() == {}
            assert img.read_xmp() == {}
            assert img.read_comment() == ''
            assert img.read_icc() == b''
            assert img.read_thumbnail() == b''

