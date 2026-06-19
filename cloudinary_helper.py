import cloudinary
import cloudinary.uploader
import os
from dotenv import load_dotenv

load_dotenv()

# I-configure ang Cloudinary gamit ang credentials mula sa .env
cloudinary.config(
    cloud_name=os.environ.get("CLOUDINARY_CLOUD_NAME"),
    api_key=os.environ.get("CLOUDINARY_API_KEY"),
    api_secret=os.environ.get("CLOUDINARY_API_SECRET"),
    secure=True
)


def upload_photo(file, folder="fcci_uploads"):
    """
    Mag-upload ng image file papunta sa Cloudinary.

    Params:
        file     — yung request.files object (hal. request.files["photo"])
        folder   — subfolder sa Cloudinary para maayos ang mga files
                   (hal. "fcci_member_photos", "fcci_receipts", atbp.)

    Returns:
        secure_url — ang public URL ng na-upload na image
                     (ito ang ise-save sa database, hindi filename)
        None       — kung walang file o may error
    """
    if not file or file.filename == "":
        return None

    try:
        result = cloudinary.uploader.upload(
            file,
            folder=folder,
            resource_type="image"
        )
        return result["secure_url"]

    except Exception as e:
        print(f"[CLOUDINARY] Upload error: {e}")
        return None


def upload_file(file, folder="fcci_uploads", resource_type="auto"):
    """
    Para sa non-image files (hal. PDF receipts).
    Gumagamit ng resource_type='auto' para awtomatiko itong
    ma-detect ng Cloudinary.
    """
    if not file or file.filename == "":
        return None

    try:
        result = cloudinary.uploader.upload(
            file,
            folder=folder,
            resource_type=resource_type
        )
        return result["secure_url"]

    except Exception as e:
        print(f"[CLOUDINARY] Upload error: {e}")
        return None


def delete_photo(public_id):
    """
    Para mag-delete ng file sa Cloudinary (optional,
    ginagamit kapag nag-delete ng member o nag-update ng photo).
    """
    try:
        cloudinary.uploader.destroy(public_id)
    except Exception as e:
        print(f"[CLOUDINARY] Delete error: {e}")
