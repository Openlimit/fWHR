import math

from matplotlib.pyplot import imshow
from PIL import Image, ImageDraw
import pylab
import urllib.request

import face_recognition


def load_image(path, url=False, imagename=''):
    if not url:
        return face_recognition.load_image_file(path)
    else:
        if path[-3:] == 'jpg' or path[-3:] == 'peg':
            urllib.request.urlretrieve(path, 'image/' + imagename)
            return face_recognition.load_image_file('image/' + imagename)
        elif path[-3:] == 'png':
            urllib.request.urlretrieve(path, 'image/' + imagename)
            return face_recognition.load_image_file('image/' + imagename)
        elif path[-3:] == 'JPG':
            urllib.request.urlretrieve(path, 'image/' + imagename)
            return face_recognition.load_image_file('image/' + imagename)
        else:
            print("Unknown image type")


def get_face_points(points, method='average', top='eyebrow'):
    width_left, width_right = points[0], points[16]

    if top == 'eyebrow':
        top_left = points[18]
        top_right = points[25]

    elif top == 'eyelid':
        top_left = points[37]
        top_right = points[43]

    else:
        raise ValueError('Invalid top point, use either "eyebrow" or "eyelid"')

    bottom_left, bottom_right = points[50], points[52]

    if method == 'left':
        coords = (width_left[0], width_right[0], top_left[1], bottom_left[1])

    elif method == 'right':
        coords = (width_left[0], width_right[0], top_right[1], bottom_right[1])

    else:
        top_average = int((top_left[1] + top_right[1]) / 2)
        bottom_average = int((bottom_left[1] + bottom_right[1]) / 2)
        coords = (width_left[0], width_right[0], top_average, bottom_average)

    ## Move the line just a little above the top of the eye to the eyelid
    if top == 'eyelid':
        coords = (coords[0], coords[1], coords[2] - 4, coords[3])

    return {'top_left': (coords[0], coords[2]),
            'bottom_left': (coords[0], coords[3]),
            'top_right': (coords[1], coords[2]),
            'bottom_right': (coords[1], coords[3])
            }


def good_picture_check(p, debug=False):
    ## To scale for picture size
    width_im = (p[16][0] - p[0][0]) / 100

    ## Difference in height between eyes
    eye_y_l = (p[37][1] + p[41][1]) / 2.0
    eye_y_r = (p[44][1] + p[46][1]) / 2.0
    eye_dif = (eye_y_r - eye_y_l) / width_im

    ## Difference top / bottom point nose
    nose_dif = (p[30][0] - p[27][0]) / width_im

    ## Space between face-edge to eye, left vs. right
    left_space = p[36][0] - p[0][0]
    right_space = p[16][0] - p[45][0]
    space_ratio = left_space / right_space

    if debug:
        print(eye_dif, nose_dif, space_ratio)

    ## These rules are not perfect, determined by trying a bunch of "bad" pictures
    if eye_dif > 5 or nose_dif > 3.5 or space_ratio > 3:
        return False
    else:
        return True


def FWHR_calc(corners):
    width = corners['top_right'][0] - corners['top_left'][0]
    height = corners['bottom_left'][1] - corners['top_left'][1]
    return float(width) / float(height)


def show_box(image, corners):
    pil_image = Image.fromarray(image)
    w, h = pil_image.size

    ## Automatically determine width of the line depending on size of picture
    line_width = math.ceil(h / 100)

    d = ImageDraw.Draw(pil_image)
    d.line([corners['bottom_left'], corners['top_left']], width=line_width)
    d.line([corners['bottom_left'], corners['bottom_right']], width=line_width)
    d.line([corners['top_left'], corners['top_right']], width=line_width)
    d.line([corners['top_right'], corners['bottom_right']], width=line_width)

    imshow(pil_image)
    pylab.show()


def get_fwhr(image_path, url=False, show=True, method='average', top='eyelid', imagename=''):
    image = load_image(image_path, url, imagename=imagename)
    landmarks = face_recognition.api._raw_face_landmarks(image)
    landmarks_as_tuples = [(p.x, p.y) for p in landmarks[0].parts()]

    if good_picture_check(landmarks_as_tuples):
        corners = get_face_points(landmarks_as_tuples, method=method, top=top)
        fwh_ratio = FWHR_calc(corners)

        if show:
            print('The Facial-Width-Height ratio is: {}'.format(fwh_ratio))
            show_box(image, corners)
        else:
            return fwh_ratio
    else:
        print("Picture is not suitable to calculate fwhr.")
        if show:
            imshow(image)
            pylab.show()
        return None


if __name__ == '__main__':
    image_url = 'http://person.sac.net.cn/photo/images/20080102/registrationRpInfo/136386186695846063.jpg'
#     # get_fwhr(obama_url, url=True, top = 'eyelid')
#     # get_fwhr(obama_url, url=True, top = 'eyebrow')
    print(get_fwhr(image_url, url=True, show=False, imagename='tmp'))
