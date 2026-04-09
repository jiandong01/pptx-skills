#!/usr/bin/env python3
"""Replace the logo image in a pptx template.

The logo lives in a single location: the slide master's image relationship (rId17).
All content layouts inherit it automatically. Structural layouts (cover, TOC,
section header) suppress master shapes and are unaffected.

Usage:
    python3 scripts/replace_logo.py new_logo.png [-t template.pptx]
    python3 scripts/replace_logo.py new_logo.png -t examples/default.pptx
"""
import argparse, os, re, shutil, zipfile


def replace_logo(template_path: str, new_logo_path: str):
    # Detect the existing logo's rId in the master .rels
    tmp = template_path + ".tmp"
    shutil.copy2(template_path, tmp)

    with zipfile.ZipFile(tmp, 'r') as z:
        contents = {n: z.read(n) for n in z.namelist()}

    # Find the image rId in the master .rels
    master_rels = contents["ppt/slideMasters/_rels/slideMaster1.xml.rels"].decode()
    m = re.search(
        r'Id="([^"]+)"[^>]*Type="[^"]*relationships/image[^"]*"[^>]*Target="([^"]+)"',
        master_rels
    )
    if not m:
        raise ValueError("No image relationship found in slide master .rels")

    rid, old_target = m.group(1), m.group(2)
    # Resolve target path inside zip: "../media/image6.png" → "ppt/media/image6.png"
    old_media_path = "ppt/" + old_target.lstrip("../")

    # Determine new media filename and path
    _, ext = os.path.splitext(new_logo_path)
    new_media_name = f"logo{ext}"
    new_media_path = f"ppt/media/{new_media_name}"

    # Replace media file
    with open(new_logo_path, 'rb') as f:
        contents[new_media_path] = f.read()

    # Remove old media if it's no longer referenced elsewhere
    if old_media_path != new_media_path:
        del contents[old_media_path]

    # Update master .rels to point to new file
    new_target = f"../media/{new_media_name}"
    master_rels_updated = re.sub(
        rf'(Id="{rid}"[^>]*Target="){re.escape(old_target)}"',
        rf'\g<1>{new_target}"',
        master_rels
    )
    contents["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = master_rels_updated.encode()

    # Update [Content_Types].xml if new extension is different
    old_ext = os.path.splitext(old_target)[-1].lower()
    new_ext = ext.lower()
    if old_ext != new_ext:
        ct_xml = contents["[Content_Types].xml"].decode()
        ext_map = {'.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
                   '.svg': 'image/svg+xml', '.gif': 'image/gif'}
        new_ct = ext_map.get(new_ext, f'image/{new_ext.lstrip(".")}')
        if f'Extension="{new_ext.lstrip(".")}"' not in ct_xml:
            ct_xml = ct_xml.replace(
                '</Types>',
                f'<Default Extension="{new_ext.lstrip(".")}" ContentType="{new_ct}"/></Types>'
            )
        contents["[Content_Types].xml"] = ct_xml.encode()

    # Write out
    os.remove(tmp)
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in contents.items():
            zout.writestr(name, data)
    os.replace(tmp, template_path)
    print(f"Replaced logo: {old_target} → {new_target} in {template_path}")


def main():
    parser = argparse.ArgumentParser(description='Replace logo in pptx template')
    parser.add_argument('logo', help='New logo image file (PNG/JPG/SVG)')
    parser.add_argument('-t', '--template', default='examples/default.pptx',
                        help='Template pptx to modify (default: examples/default.pptx)')
    args = parser.parse_args()
    if not os.path.exists(args.logo):
        print(f"Error: {args.logo} not found"); return
    if not os.path.exists(args.template):
        print(f"Error: {args.template} not found"); return
    replace_logo(args.template, args.logo)


if __name__ == '__main__':
    main()
