import re
import os
import shutil
from pdb import set_trace

import fitz
import win32com.client as win32
import pywintypes
import traceback
from tqdm import tqdm
from xml.etree import ElementTree

### OPTIONS ###
PROCESS_RECYCLE_BIN = False
ASSETS_DIR = "assets"
OUTPUT_FOLDER = "OneNoteExport"
KEEP_INTERMEDIATE = True
FIX_DIMENSIONS = True
FIX_BACKSLASH = True
FIX_HEADER = True
REMOVE_NBSP = True
ADD_CONVERSION_TAG = True
TARGET_NB = "Fliegerei"  # If None, all notebooks will be converted

# If the user uses OneDrive to sync the Desktop, use the according directory,
# otherwise use the regular desktop.
ONEDRIVE_DIR = os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop")
if os.path.exists(ONEDRIVE_DIR):
    OUTPUT_DIR = os.path.join(ONEDRIVE_DIR, OUTPUT_FOLDER)
else:
    OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "OneNoteExport")
LOGFILE = "onenote_to_markdown.log"  # Set to None to disable logging
ILLEGAL_WIN_FILENAMES = "CON, PRN, AUX, NUL, COM1, COM2, COM3, COM4, COM5, COM6, COM7, COM8, COM9, LPT1, LPT2, LPT3, LPT4, LPT5, LPT6, LPT7, LPT8, LPT9"


def log(message, tqdm=None):
    if tqdm:
        # Update progress bar description, leave 50 columns space for the bar itself
        width = os.get_terminal_size().columns
        tqdm.set_description(truncate(message, width - 50))
    else:
        print(message)
    if LOGFILE is not None:
        with open(LOGFILE, "a") as lf:
            lf.write("%s\n" % message)


def truncate(string, length, ellipsis="..."):
    """Truncates a string to a certain `length`,
    by putting an `ellipsis` in the middle."""
    if len(string) <= length:
        return string
    left = length // 2
    right = length - left
    return (
        string[: left - len(ellipsis) // 2]
        + ellipsis
        + string[-right + len(ellipsis) // 2 + len(ellipsis) % 2 :]
    )


def safe_str(name):
    """Replaces illegal characters in filenames.
    Characters according to https://stackoverflow.com/a/31976060"""
    if name in ILLEGAL_WIN_FILENAMES.split(", "):
        return name + "_note"
    return re.sub(r"[<>:\"\\/|?*]", "_", name.strip())


def replace_whitespace(name):
    """Replaces whitespace with an underscore."""
    return name.replace(" ", "_")


def extract_pdf_pictures(pdf_path, assets_path, page_name, tqdm):
    os.makedirs(assets_path, exist_ok=True)
    image_names = []
    doc = fitz.open(pdf_path)
    img_num = 0
    for i in range(len(doc)):
        for img in doc.get_page_images(i):
            xref = img[0]
            pix = fitz.Pixmap(doc, xref)
            png_name = f"{replace_whitespace(page_name)}_{str(img_num).zfill(3)}.png"
            png_path = os.path.join(assets_path, png_name)
            log("Writing png: %s" % png_path, tqdm)
            if pix.n < 5:
                pix.save(png_path)
            else:
                pix1 = fitz.Pixmap(fitz.csRGB, pix)
                pix1.save(png_path)
                pix1 = None
            pix = None
            image_names.append(png_name)
            img_num += 1
    return image_names


def clean_pandoc_result(md_path, image_names):
    tmp_path = md_path + ".tmp"
    with open(md_path, "r", encoding="utf-8") as f_md, open(
        tmp_path, "w", encoding="utf-8"
    ) as f_tmp:
        # body_md: str = f_md.read()
        lines = f_md.readlines()
        if FIX_HEADER:
            lines[0] = "# " + lines[0]
            lines[2] = lines[2].strip() + " " + lines[4]
            # lines[3] = f"parent::[[{parent}]]"
            if ADD_CONVERSION_TAG:
                lines[3] = "#FromOneNote\n"
            lines[4] = "___"
        body_md = "".join(lines)
        # set_trace()
        body_md = fix_image_names(body_md, image_names)
        if FIX_DIMENSIONS:
            body_md = convert_image_dimensions_obsidian(body_md)
        if FIX_BACKSLASH:
            body_md = remove_backslashes(body_md)
        if REMOVE_NBSP:
            body_md = remove_nsbp(body_md)
        body_md = fix_blank_lines(body_md)
        body_md = convert_crlf_to_lf(body_md)
        f_tmp.write(body_md)
    shutil.move(tmp_path, md_path)


def fix_image_names(body_md, image_names):
    for i, name in enumerate(image_names):
        body_md = re.sub(r"media\/image" + str(i + 1) + r"\.[a-zA-Z]+", name, body_md)
    return body_md


def convert_image_dimensions_obsidian(md: str, ppi: int = 96):
    def replace(result):
        alt, name, width_in, height_in = (
            result.group(1),
            result.group(2),
            result.group(3),
            result.group(4),
        )
        width_px, height_px = int(float(width_in) * ppi), int(float(height_in) * ppi)
        return f"![{alt}|{width_px}x{height_px}]({name})"

    return re.sub(
        r'\!\[(.*)\]\((.*)\){width="(\d+\.\d+)in" height="(\d+\.\d+)in"}', replace, md
    )


def remove_backslashes(body_md):
    return body_md.replace('\\"', '"').replace("\\'", "'").replace("\\...", "...")
    # TODO: replace all backslashes except before alphanumeric characters


def remove_nsbp(body_md):
    return body_md.replace("\xa0", "")


def convert_crlf_to_lf(body_md):
    return re.sub(r"\r\n", r"\n", body_md)


def fix_blank_lines(body_md):
    return re.sub(
        r"\r*\n[ \t]*>?\r*\n(?![|])", r"\n", body_md
    )  # Remove ALL double blank lines except before tables
    # return re.sub(r"\r*\n\r*\n([ \t]*)- ", r"\n\1- ", body_md)  # Remove only blank lines in lists


def handle_page(onenote, elem, path, i, tqdm=None):
    full_path = os.path.join(OUTPUT_DIR, path)
    os.makedirs(full_path, exist_ok=True)
    path_assets = os.path.join(full_path, ASSETS_DIR)
    safe_name = safe_str(str(i).zfill(3) + " " + elem.attrib["name"])
    safe_path = os.path.join(full_path, safe_name)
    path_docx = safe_path + ".docx"
    path_pdf = safe_path + ".pdf"
    path_md = safe_path + ".md"
    # Remove temp files
    if not KEEP_INTERMEDIATE:
        try:
            os.remove(path_docx)
            os.remove(path_pdf)
        except OSError:
            pass
    try:
        # Create docx
        if not os.path.exists(path_docx):
            onenote.Publish(elem.attrib["ID"], path_docx, win32.constants.pfWord, "")
        # Convert docx to markdown
        log("Generating markdown: %s" % path_md, tqdm)
        os.system(
            f'pandoc.exe -i "{path_docx}" -o "{path_md}" -t markdown-simple_tables-multiline_tables-grid_tables --wrap=none'
        )
        # Create pdf (for the picture assets)
        if not os.path.exists(path_pdf):
            onenote.Publish(elem.attrib["ID"], path_pdf, 3, "")
        # Output picture assets to folder
        image_names = extract_pdf_pictures(path_pdf, path_assets, safe_name, tqdm)
        # Replace image names in markdown file
        clean_pandoc_result(path_md, image_names)
    except pywintypes.com_error as e:
        log("!!WARNING!! Page Failed: %s" % path_md)
    # Clean up docx, html
    if not KEEP_INTERMEDIATE:
        try:
            os.remove(path_docx)
            os.remove(path_pdf)
        except OSError:
            pass
    # Delete PDF in any way!
    # try:
    #     os.remove(path_pdf)
    # except OSError:
    #     pass


def handle_element(onenote, elem, path="", i=0, tqdm=None, last_elem=None):
    if elem.tag.endswith("Notebook"):
        hier2 = onenote.GetHierarchy(elem.attrib["ID"], win32.constants.hsChildren, "")
        for i, c2 in enumerate(ElementTree.fromstring(hier2)):
            handle_element(
                onenote, c2, os.path.join(path, safe_str(elem.attrib["name"])), i, tqdm
            )
    elif elem.tag.endswith("Section"):
        hier2 = onenote.GetHierarchy(elem.attrib["ID"], win32.constants.hsPages, "")
        for i, c2 in enumerate(ElementTree.fromstring(hier2)):
            handle_element(
                onenote, c2, os.path.join(path, safe_str(elem.attrib["name"])), i, tqdm
            )
    elif elem.tag.endswith("SectionGroup") and (
        not elem.attrib["name"].startswith("OneNote_RecycleBin") or PROCESS_RECYCLE_BIN
    ):
        hier2 = onenote.GetHierarchy(elem.attrib["ID"], win32.constants.hsSections, "")
        for i, c2 in enumerate(ElementTree.fromstring(hier2)):
            handle_element(
                onenote, c2, os.path.join(path, safe_str(elem.attrib["name"])), i, tqdm
            )
    elif elem.tag.endswith("Page"):
        try:
            if elem.attrib["isSubPage"] == "true":
                # Create parent directory, but how to find out parent?
                handle_page(onenote, elem, path, i, tqdm)
        except KeyError:
            handle_page(onenote, elem, path, i, tqdm)


if __name__ == "__main__":
    try:
        onenote = win32.gencache.EnsureDispatch("OneNote.Application.12")

        hier = onenote.GetHierarchy("", win32.constants.hsNotebooks, "")

        root = ElementTree.fromstring(hier)
        pbar = tqdm(root)
        for child in pbar:
            if TARGET_NB and TARGET_NB != child.attrib["name"]:
                continue
            pbar.set_description(f"Processing Notebook {child.attrib['name']}")
            handle_element(onenote, child, tqdm=pbar)

    except pywintypes.com_error as e:
        traceback.print_exc()
        log(
            "!!!Error!!! Hint: Make sure OneNote is open first; run OneNote and PowerShell as Administrator!"
        )
