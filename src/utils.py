import io
from PIL import Image

def load_image(image_file):
    """
    讀取上傳的圖片並轉換為 PIL Image 物件
    """
    try:
        image = Image.open(image_file)
        # 修正 EXIF 方向
        try:
            from PIL import ImageOps
            image = ImageOps.exif_transpose(image)
        except Exception:
            pass
        return image
    except Exception as e:
        print(f"Error loading image: {e}")
        return None

def compress_image(image, max_size=(1024, 1024), quality=85):
    """
    壓縮圖片以減少 Word 檔案大小
    """
    img_copy = image.copy()
    img_copy.thumbnail(max_size, Image.Resampling.LANCZOS)
    
    output_buffer = io.BytesIO()
    # 轉換為 RGB
    if img_copy.mode in ('RGBA', 'P'):
        img_copy = img_copy.convert('RGB')
        
    img_copy.save(output_buffer, format='JPEG', quality=quality)
    return output_buffer

def get_image_date(image):
    """
    嘗試從 EXIF 讀取拍攝日期
    """
    try:
        from PIL.ExifTags import TAGS
        exif = image._getexif()
        if not exif:
            return None
            
        for tag, value in exif.items():
            decoded = TAGS.get(tag, tag)
            if decoded == "DateTimeOriginal":
                # Format: YYYY:MM:DD HH:MM:SS
                # Return standardized string YYYY-MM-DD
                return parts
    except Exception:
        return None
    return None

def crop_to_ratio(image, target_ratio=1.47):
    """
    Center crop the image to the target aspect ratio.
    Default ratio 1.47 is approx 14.4/9.8
    """
    img_w, img_h = image.size
    current_ratio = img_w / img_h
    
    if current_ratio > target_ratio:
        # Too wide, crop width
        new_w = int(img_h * target_ratio)
        offset = (img_w - new_w) // 2
        return image.crop((offset, 0, offset + new_w, img_h))
    else:
        return image.crop((0, offset, img_w, offset + new_h))

def resize_with_padding(image, target_ratio=1.47, bg_color=(255, 255, 255)):
    """
    Resize image to fit within target aspect ratio, padding with background color.
    (Letterboxing)
    """
    img_w, img_h = image.size
    current_ratio = img_w / img_h
    
    # Target dimensions (base on width=1000px for good resolution)
    base_w = 1000
    base_h = int(base_w / target_ratio)
    
    # Create background
    new_img = Image.new("RGB", (base_w, base_h), bg_color)
    
    if current_ratio > target_ratio:
        # Image is wider (relatively) -> Fit Width
        # Resized Width = base_w
        # Resized Height = base_w / current_ratio
        w = base_w
        h = int(base_w / current_ratio)
        x = 0
        y = (base_h - h) // 2
    else:
        # Image is taller (relatively) -> Fit Height
        # Resized Height = base_h
        # Resized Width = base_h * current_ratio
        h = base_h
        w = int(base_h * current_ratio)
        x = (base_w - w) // 2
        y = 0
        
    resized_content = image.resize((w, h), Image.Resampling.LANCZOS)
    new_img.paste(resized_content, (x, y))
    
    return new_img
