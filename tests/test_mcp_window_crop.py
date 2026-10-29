"""Window cropping must preserve complete content and its screen-coordinate origin."""

from PIL import Image, ImageDraw

from easy_uiauto.visual import crop_window_image


def _window(busy=True, caption=True):
    image = Image.new("RGB", (480, 320), "white")
    draw = ImageDraw.Draw(image)
    if busy:
        for x in range(0, 480, 7):
            draw.line((x, 0, x, 319), fill=(x % 255, 70, 150))
        draw.text((1, 1), "Underlying window content", fill="black")
    draw.rectangle((16, 16, 463, 303), fill=(0, 61, 116))
    draw.rectangle((17, 17, 462, 302), fill="white")
    if caption:
        draw.rectangle((17, 17, 462, 40), fill=(209, 224, 239))
        draw.text((24, 23), "Window", fill="black")
        draw.text((445, 23), "X", fill="black")
    draw.rectangle((33, 70, 440, 280), outline=(180, 180, 180))
    return image


def test_busy_exterior_crops_complete_visible_window():
    image = _window()
    result = crop_window_image(image)
    assert result["cropped"] is True
    assert result["rect_in_input"] == dict(left=16, top=16, right=464, bottom=304,
                                          width=448, height=288)
    assert result["evidence"] == "window-frame-and-titlebar"
    assert result["image"].tobytes() == image.crop((16, 16, 464, 304)).tobytes()


def test_screen_mapping_retains_negative_monitor_origin():
    result = crop_window_image(_window(), dict(left=-900, top=120, width=480, height=320))
    assert result["screen_rect"] == dict(left=-884, top=136, right=-436, bottom=424,
                                        width=448, height=288)
    assert result["offset"] == dict(x=16, y=16)


def test_similar_coloured_shadow_does_not_add_an_exterior_scanline():
    image = _window()
    draw = ImageDraw.Draw(image)
    draw.rectangle((20, 0, 250, 15), fill=(10, 50, 100))
    draw.line((464, 17, 464, 302), fill=(4, 73, 129))
    draw.line((17, 304, 462, 304), fill=(4, 73, 129))
    result = crop_window_image(image)
    assert result["rect_in_input"] == dict(left=16, top=16, right=464, bottom=304,
                                          width=448, height=288)
    assert result["image"].tobytes() == image.crop((16, 16, 464, 304)).tobytes()


def test_already_cropped_window_keeps_title_bar_and_white_client_area():
    image = _window().crop((16, 16, 464, 304))
    result = crop_window_image(image)
    assert result["cropped"] is False
    assert result["image"] is image


def test_busy_exterior_without_caption_evidence_is_ambiguous():
    image = _window(caption=False)
    result = crop_window_image(image)
    assert result["cropped"] is False


def test_small_control_is_not_selected_as_the_window():
    image = Image.new("RGB", (480, 320), "white")
    ImageDraw.Draw(image).rectangle((160, 140, 320, 180), outline="black", fill="gray")
    assert crop_window_image(image)["cropped"] is False


def test_padding_expands_from_actual_visible_frame():
    result = crop_window_image(_window(), padding=2)
    assert result["rect_in_input"] == dict(left=14, top=14, right=466, bottom=306,
                                          width=452, height=292)
