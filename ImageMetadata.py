#Extract Image Metadata by hscho

import os, folium, csv
from PIL import Image
from PIL.ExifTags import TAGS, GPSTAGS

def getPath() :
    Path = os.getcwd()
    return Path

def search(directory):
    try:
        filenames = os.listdir(directory)
        for filename in filenames :
            full_filename = os.path.join(directory, filename)
            if os.path.isdir(full_filename) :
                search(full_filename)
            else:
                ext = os.path.splitext(full_filename)[-1]
                if ext == '.jpg' or ext == '.jpeg' or ext == 'JPG' or ext == 'JPEG' :
                    imageList.append(full_filename)
        return(imageList)
    except PermissionError:
        pass

def getExif(imagefile) :
    image = Image.open(imagefile)
    img = image._getexif()
    return img

def labeledExif(exif) :
    labeled = {}
    for (key, val) in exif.items() :
        labeled[TAGS.get(key)] = val
    return labeled

def get_geotagging(exif):
    if not exif:
        pass
    geotagging = {}
    for (idx, tag) in TAGS.items():
        if tag == 'GPSInfo':
            if idx not in exif:
                pass
            for (key, val) in GPSTAGS.items():
                if key in exif[idx]:
                    geotagging[val] = exif[idx][key]
    return geotagging

def get_decimal_from_dms(dms, ref):
    degrees = dms[0][0] / dms[0][1]
    minutes = dms[1][0] / dms[1][1] / 60.0
    seconds = dms[2][0] / dms[2][1] / 3600.0
    if ref in ['S', 'W']:
        degrees = -degrees
        minutes = -minutes
        seconds = -seconds
    return round(degrees + minutes + seconds, 5)

def get_coordinates(geotags):
    lat = get_decimal_from_dms(geotags['GPSLatitude'], geotags['GPSLatitudeRef'])
    lon = get_decimal_from_dms(geotags['GPSLongitude'], geotags['GPSLongitudeRef'])
    return lat, lon

def CalculateFileSize(file_size):
    if file_size >= (1024 * 1024):
        file_size = file_size / 1024 / 1024
    else:
        file_size = file_size / 1024
    return file_size

def CreateThumbnail(image):
    img = image.open(image)
    size = (64, 64)
    img.thumbnail(size)
    img.save('C:/Output/thumbnail/%s_thumbnail.jpg'%image)

print("Processing...", "\n")

imageList = []

if not os.path.isdir("C:/Output"):
    os.mkdir("C:" + "/" + "Output" + "/")
if not os.path.isdir("C:/Output/thumbnail"):
    os.mkdir("C:" + "/" + "Output" + "/" + "thumbnail" + "/")
search(getPath())

with open('C:/Output/ImageMetadata.csv', 'w', encoding='euc-kr', newline = '') as output:
    writer = csv.writer(output)
    writer.writerow(['번호', "파일명", "찍은 날짜", "너비", "높이", "카메라 제조업체", "카메라 모델", "노출 시간",
                    "사진 크기", "파일 경로", "촬영 장소"])
    output.close()

count = 1
for x in imageList :
    img = Image.open(x)
    width, height = img.size

    thumb_size = (50, 50)
    img.thumbnail(thumb_size)
    img.save('C:/Output/thumbnail/%s_thumbnail.jpg'%x.split("\\")[-1])

    size = os.path.getsize(x)
    img_size = CalculateFileSize(size)
    if size >= (1024 * 1024):
        img_size = str(round(img_size, 2)) + "MB"
    else:
        img_size = str(round(img_size, 2)) + "KB"

    try:
        tag = labeledExif(getExif(x))

        filename = x.split('\\')

        geotag = get_geotagging(getExif(x))
        coordinates = get_coordinates(geotag)

        map = folium.Map(location = [coordinates[0], coordinates[1]], zoom_start = 35)
        folium.Marker([coordinates[0], coordinates[1]], icon = folium.Icon(color='red',icon='star')).add_to(map)
        map.save('C:/Output/%s_map.html' %filename[-1])

        row = [count, filename[-1], tag.get('DateTimeOriginal'), tag.get('ExifImageWidth'), tag.get('ExifImageHeight'),
                tag.get('Make'), tag.get('Model'), tag.get('ExposureTime'), img_size, x, 'file:///C:/Output/%s_map.html'%filename[-1]]

    except:
        row = [count, filename[-1], tag.get('DateTimeOriginal'), tag.get('ExifImageWidth'), tag.get('ExifImageHeight'),
            tag.get('Make'), tag.get('Model'), tag.get('ExposureTime'), img_size, x]

    with open('C:/Output/ImageMetadata.csv', 'a', newline = '') as output:
        writer = csv.writer(output)
        try:
            writer.writerow(row)
        except AttributeError:
            pass

    output.close()
    count += 1

print("Processing Done")
