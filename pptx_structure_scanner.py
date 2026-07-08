import zipfile
import os
import json
import re
import html
import argparse
import hashlib
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    from PIL import Image
except ImportError:
    Image = None


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


EMU_PER_PIXEL_APPROX = 9525


def safe_xml_parse(text):
    try:
        return ET.fromstring(text)
    except Exception:
        return None


def get_local_name(tag):
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def read_zip_text(zf, path):
    try:
        return zf.read(path).decode("utf-8", errors="replace")
    except Exception:
        return ""


def sha1_bytes(data):
    return hashlib.sha1(data).hexdigest()


def get_image_info(data, filename):
    info = {
        "width": None,
        "height": None,
        "format": None,
    }

    if Image is None:
        return info

    try:
        from io import BytesIO
        img = Image.open(BytesIO(data))
        info["width"] = img.width
        info["height"] = img.height
        info["format"] = img.format
    except Exception:
        pass

    return info


def normalize_target(base_rels_path, target):
    """
    Converts relationship Target into a ppt/... path when possible.
    Example:
    rel path: ppt/slides/_rels/slide1.xml.rels
    target: ../media/image1.png
    result: ppt/media/image1.png
    """
    if target.startswith("/"):
        target = target.lstrip("/")
        return target

    rel_dir = Path(base_rels_path).parent
    # remove "_rels" folder
    if rel_dir.name == "_rels":
        parent_dir = rel_dir.parent
    else:
        parent_dir = rel_dir

    combined = parent_dir / target
    normalized = os.path.normpath(str(combined)).replace("\\", "/")
    return normalized


def parent_xml_from_rels(rels_path):
    """
    ppt/slides/_rels/slide1.xml.rels -> ppt/slides/slide1.xml
    """
    return rels_path.replace("/_rels/", "/").replace(".rels", "")


def get_slide_number_from_path(path):
    m = re.search(r"ppt/slides/slide(\d+)\.xml$", path)
    if m:
        return int(m.group(1))
    return None


def get_context_for_blip(xml_root, rel_id):
    """
    Find image blips using r:embed rel id, then walk upward to get picture/shape context.
    """
    traces = []

    parent_map = {child: parent for parent in xml_root.iter() for child in parent}

    for elem in xml_root.iter():
        if get_local_name(elem.tag) != "blip":
            continue

        embed = elem.attrib.get(f"{{{NS['r']}}}embed") or elem.attrib.get("embed")
        link = elem.attrib.get(f"{{{NS['r']}}}link") or elem.attrib.get("link")

        if embed != rel_id and link != rel_id:
            continue

        context = {
            "rel_id": rel_id,
            "blip_tag": elem.tag,
            "shape_type": None,
            "shape_name": None,
            "shape_id": None,
            "x": None,
            "y": None,
            "cx": None,
            "cy": None,
            "x_px": None,
            "y_px": None,
            "w_px": None,
            "h_px": None,
            "xml_snippet": "",
        }

        current = elem
        container = None

        while current is not None:
            local = get_local_name(current.tag)
            if local in ["pic", "sp", "graphicFrame", "bg", "grpSp"]:
                container = current
                context["shape_type"] = local
                break
            current = parent_map.get(current)

        if container is not None:
            # shape name/id
            for cNvPr in container.iter():
                if get_local_name(cNvPr.tag) == "cNvPr":
                    context["shape_name"] = cNvPr.attrib.get("name")
                    context["shape_id"] = cNvPr.attrib.get("id")
                    break

            # position and size
            for off in container.iter():
                if get_local_name(off.tag) == "off":
                    context["x"] = off.attrib.get("x")
                    context["y"] = off.attrib.get("y")
                    break

            for ext in container.iter():
                if get_local_name(ext.tag) == "ext":
                    context["cx"] = ext.attrib.get("cx")
                    context["cy"] = ext.attrib.get("cy")
                    break

            def emu_to_px(v):
                try:
                    return round(int(v) / EMU_PER_PIXEL_APPROX, 2)
                except Exception:
                    return None

            context["x_px"] = emu_to_px(context["x"])
            context["y_px"] = emu_to_px(context["y"])
            context["w_px"] = emu_to_px(context["cx"])
            context["h_px"] = emu_to_px(context["cy"])

            raw = ET.tostring(container, encoding="unicode", method="xml")
            context["xml_snippet"] = raw[:3000]

        traces.append(context)

    return traces


def scan_pptx(pptx_path):
    pptx_path = Path(pptx_path)

    if not pptx_path.exists():
        raise FileNotFoundError(f"File not found: {pptx_path}")

    report = {
        "file": str(pptx_path),
        "file_size_bytes": pptx_path.stat().st_size,
        "media": {},
        "relationships": [],
        "slides": {},
        "usage_summary": {},
        "suspicious_candidates": [],
        "all_zip_files": [],
        "notes": [],
    }

    with zipfile.ZipFile(pptx_path, "r") as zf:
        names = zf.namelist()
        report["all_zip_files"] = names

        media_files = [
            n for n in names
            if n.lower().startswith("ppt/media/")
            and re.search(r"\.(png|jpe?g|gif|webp|svg|emf|wmf)$", n, re.I)
        ]

        rel_files = [
            n for n in names
            if n.lower().endswith(".xml.rels")
        ]

        xml_files = [
            n for n in names
            if n.lower().endswith(".xml")
        ]

        # Collect media info
        for media_path in media_files:
            data = zf.read(media_path)
            filename = media_path.split("/")[-1]
            ext = filename.split(".")[-1].lower() if "." in filename else ""

            image_info = get_image_info(data, filename)

            report["media"][media_path] = {
                "path": media_path,
                "filename": filename,
                "extension": ext,
                "size_bytes": len(data),
                "sha1": sha1_bytes(data),
                "width": image_info["width"],
                "height": image_info["height"],
                "format": image_info["format"],
                "references": [],
            }

        # Parse relationships
        for rel_path in rel_files:
            rel_text = read_zip_text(zf, rel_path)
            rel_root = safe_xml_parse(rel_text)

            if rel_root is None:
                continue

            parent_xml = parent_xml_from_rels(rel_path)

            for rel in rel_root:
                if get_local_name(rel.tag) != "Relationship":
                    continue

                rel_id = rel.attrib.get("Id")
                rel_type = rel.attrib.get("Type", "")
                target = rel.attrib.get("Target", "")
                target_mode = rel.attrib.get("TargetMode", "")

                normalized_target = normalize_target(rel_path, target)

                item = {
                    "rels_file": rel_path,
                    "parent_xml": parent_xml,
                    "rel_id": rel_id,
                    "type": rel_type,
                    "target": target,
                    "normalized_target": normalized_target,
                    "target_mode": target_mode,
                    "is_image": normalized_target in report["media"],
                    "slide_number": get_slide_number_from_path(parent_xml),
                    "xml_traces": [],
                }

                if normalized_target in report["media"] and parent_xml in names:
                    xml_text = read_zip_text(zf, parent_xml)
                    xml_root = safe_xml_parse(xml_text)

                    if xml_root is not None:
                        traces = get_context_for_blip(xml_root, rel_id)
                        item["xml_traces"] = traces

                report["relationships"].append(item)

                if normalized_target in report["media"]:
                    report["media"][normalized_target]["references"].append(item)

        # Slide overview
        slide_xml_files = sorted(
            [n for n in xml_files if re.match(r"ppt/slides/slide\d+\.xml$", n)],
            key=lambda x: get_slide_number_from_path(x) or 0
        )

        for slide_path in slide_xml_files:
            slide_no = get_slide_number_from_path(slide_path)
            slide_text = read_zip_text(zf, slide_path)
            slide_root = safe_xml_parse(slide_text)

            slide_info = {
                "slide_path": slide_path,
                "slide_number": slide_no,
                "image_refs": [],
                "shape_count": 0,
                "pic_count": 0,
                "blip_count": 0,
                "raw_xml_preview": slide_text[:2000],
            }

            if slide_root is not None:
                for elem in slide_root.iter():
                    local = get_local_name(elem.tag)
                    if local == "sp":
                        slide_info["shape_count"] += 1
                    elif local == "pic":
                        slide_info["pic_count"] += 1
                    elif local == "blip":
                        slide_info["blip_count"] += 1
                        embed = elem.attrib.get(f"{{{NS['r']}}}embed") or elem.attrib.get("embed")
                        link = elem.attrib.get(f"{{{NS['r']}}}link") or elem.attrib.get("link")
                        slide_info["image_refs"].append(embed or link)

            report["slides"][str(slide_no)] = slide_info

        # Usage summary and suspicious candidates
        for media_path, media in report["media"].items():
            refs = media["references"]
            slide_refs = [r for r in refs if r["slide_number"] is not None]
            unique_slides = sorted(set(r["slide_number"] for r in slide_refs))

            total_traces = sum(len(r.get("xml_traces", [])) for r in refs)

            summary = {
                "path": media_path,
                "filename": media["filename"],
                "size_bytes": media["size_bytes"],
                "width": media["width"],
                "height": media["height"],
                "reference_count": len(refs),
                "slide_reference_count": len(unique_slides),
                "slides": unique_slides,
                "xml_trace_count": total_traces,
                "sha1": media["sha1"],
            }

            report["usage_summary"][media_path] = summary

            score = 0
            reasons = []

            if len(unique_slides) >= 2:
                score += 3
                reasons.append(f"Used on {len(unique_slides)} slides")

            if len(refs) >= 3:
                score += 2
                reasons.append(f"{len(refs)} total relationship references")

            if media["size_bytes"] <= 200000:
                score += 1
                reasons.append("Small file size")

            if media["width"] and media["height"]:
                area = media["width"] * media["height"]

                if area <= 500000:
                    score += 1
                    reasons.append("Small/medium image dimensions")

                if media["width"] >= 300 and media["height"] >= 100:
                    score += 1
                    reasons.append("Overlay-like dimensions")

            if total_traces >= 2:
                score += 1
                reasons.append("Found multiple XML placement traces")

            if score >= 4:
                report["suspicious_candidates"].append({
                    "path": media_path,
                    "filename": media["filename"],
                    "score": score,
                    "reasons": reasons,
                    "size_bytes": media["size_bytes"],
                    "width": media["width"],
                    "height": media["height"],
                    "slides": unique_slides,
                    "reference_count": len(refs),
                    "sha1": media["sha1"],
                })

        report["suspicious_candidates"].sort(
            key=lambda x: (x["score"], x["reference_count"]),
            reverse=True
        )

        if Image is None:
            report["notes"].append(
                "Pillow is not installed, so image dimensions could not be detected. Install with: pip install pillow"
            )

    return report


def escape(v):
    if v is None:
        return ""
    return html.escape(str(v))


def make_html_report(report):
    suspicious_rows = ""

    for c in report["suspicious_candidates"]:
        suspicious_rows += f"""
        <tr>
            <td><code>{escape(c["filename"])}</code></td>
            <td>{escape(c["score"])}</td>
            <td>{escape(c["size_bytes"])}</td>
            <td>{escape(c["width"])} × {escape(c["height"])}</td>
            <td>{escape(", ".join(map(str, c["slides"])))}</td>
            <td>{escape(c["reference_count"])}</td>
            <td>{escape("; ".join(c["reasons"]))}</td>
        </tr>
        """

    media_rows = ""

    for media_path, media in report["media"].items():
        summary = report["usage_summary"].get(media_path, {})
        media_rows += f"""
        <tr>
            <td><code>{escape(media["filename"])}</code></td>
            <td><code>{escape(media_path)}</code></td>
            <td>{escape(media["extension"])}</td>
            <td>{escape(media["size_bytes"])}</td>
            <td>{escape(media["width"])} × {escape(media["height"])}</td>
            <td>{escape(summary.get("reference_count"))}</td>
            <td>{escape(summary.get("slide_reference_count"))}</td>
            <td>{escape(", ".join(map(str, summary.get("slides", []))))}</td>
            <td><code>{escape(media["sha1"][:12])}</code></td>
        </tr>
        """

    relationship_rows = ""

    for r in report["relationships"]:
        if not r["is_image"]:
            continue

        traces_preview = ""
        for t in r.get("xml_traces", []):
            traces_preview += f"""
            <details>
                <summary>
                    {escape(t.get("shape_type"))}
                    | name: {escape(t.get("shape_name"))}
                    | pos: {escape(t.get("x_px"))},{escape(t.get("y_px"))}
                    | size: {escape(t.get("w_px"))}×{escape(t.get("h_px"))} px
                </summary>
                <pre>{escape(t.get("xml_snippet"))}</pre>
            </details>
            """

        relationship_rows += f"""
        <tr>
            <td>{escape(r["slide_number"])}</td>
            <td><code>{escape(r["rel_id"])}</code></td>
            <td><code>{escape(r["parent_xml"])}</code></td>
            <td><code>{escape(r["normalized_target"])}</code></td>
            <td>{traces_preview if traces_preview else "<em>No direct blip trace found</em>"}</td>
        </tr>
        """

    slide_rows = ""

    for slide_no, s in sorted(report["slides"].items(), key=lambda x: int(x[0]) if x[0] != "None" else 0):
        slide_rows += f"""
        <tr>
            <td>{escape(slide_no)}</td>
            <td><code>{escape(s["slide_path"])}</code></td>
            <td>{escape(s["shape_count"])}</td>
            <td>{escape(s["pic_count"])}</td>
            <td>{escape(s["blip_count"])}</td>
            <td><code>{escape(", ".join([str(x) for x in s["image_refs"] if x]))}</code></td>
        </tr>
        """

    notes_html = ""
    for n in report.get("notes", []):
        notes_html += f"<li>{escape(n)}</li>"

    html_doc = f"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>PPTX Structure Report</title>
<style>
    body {{
        font-family: Arial, sans-serif;
        background: #f6f7f9;
        color: #111827;
        margin: 0;
        padding: 24px;
    }}
    h1, h2 {{
        margin-bottom: 10px;
    }}
    .card {{
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        padding: 18px;
        margin-bottom: 22px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }}
    table {{
        width: 100%;
        border-collapse: collapse;
        font-size: 14px;
    }}
    th, td {{
        border: 1px solid #e5e7eb;
        padding: 8px;
        vertical-align: top;
    }}
    th {{
        background: #f3f4f6;
        text-align: left;
        position: sticky;
        top: 0;
    }}
    code {{
        background: #f3f4f6;
        padding: 2px 4px;
        border-radius: 4px;
    }}
    pre {{
        white-space: pre-wrap;
        word-break: break-word;
        background: #111827;
        color: #e5e7eb;
        padding: 12px;
        border-radius: 8px;
        max-height: 320px;
        overflow: auto;
    }}
    details {{
        margin-bottom: 8px;
    }}
    .badge {{
        display: inline-block;
        padding: 4px 8px;
        background: #eef2ff;
        color: #3730a3;
        border-radius: 999px;
        font-size: 12px;
        font-weight: bold;
    }}
</style>
</head>
<body>

<h1>PPTX Structure Report</h1>

<div class="card">
    <p><b>File:</b> <code>{escape(report["file"])}</code></p>
    <p><b>File size:</b> {escape(report["file_size_bytes"])} bytes</p>
    <p><b>Total media files:</b> {len(report["media"])}</p>
    <p><b>Suspicious candidates:</b> <span class="badge">{len(report["suspicious_candidates"])}</span></p>
    {"<ul>" + notes_html + "</ul>" if notes_html else ""}
</div>

<div class="card">
    <h2>Suspicious Watermark Candidates</h2>
    <table>
        <thead>
            <tr>
                <th>Image</th>
                <th>Score</th>
                <th>Size bytes</th>
                <th>Dimensions</th>
                <th>Slides</th>
                <th>Refs</th>
                <th>Reasons</th>
            </tr>
        </thead>
        <tbody>
            {suspicious_rows if suspicious_rows else "<tr><td colspan='7'>No suspicious candidates found by scoring.</td></tr>"}
        </tbody>
    </table>
</div>

<div class="card">
    <h2>All Media Files</h2>
    <table>
        <thead>
            <tr>
                <th>Filename</th>
                <th>Path</th>
                <th>Type</th>
                <th>Size bytes</th>
                <th>Dimensions</th>
                <th>Total refs</th>
                <th>Slide refs</th>
                <th>Slides</th>
                <th>SHA1</th>
            </tr>
        </thead>
        <tbody>
            {media_rows}
        </tbody>
    </table>
</div>

<div class="card">
    <h2>Image XML Traces</h2>
    <table>
        <thead>
            <tr>
                <th>Slide</th>
                <th>Rel ID</th>
                <th>Parent XML</th>
                <th>Image Target</th>
                <th>Trace</th>
            </tr>
        </thead>
        <tbody>
            {relationship_rows if relationship_rows else "<tr><td colspan='5'>No image relationships found.</td></tr>"}
        </tbody>
    </table>
</div>

<div class="card">
    <h2>Slide Overview</h2>
    <table>
        <thead>
            <tr>
                <th>Slide</th>
                <th>Path</th>
                <th>Shapes</th>
                <th>Pictures</th>
                <th>Blips</th>
                <th>Image Rel IDs</th>
            </tr>
        </thead>
        <tbody>
            {slide_rows}
        </tbody>
    </table>
</div>

</body>
</html>
"""
    return html_doc


def main():
    parser = argparse.ArgumentParser(description="Scan PPTX internal structure for images and XML traces.")
    parser.add_argument("pptx", help="Path to the PPTX file")
    parser.add_argument("--out", default="pptx_scan_output", help="Output folder")
    args = parser.parse_args()

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    report = scan_pptx(args.pptx)

    json_path = out_dir / "pptx_report.json"
    html_path = out_dir / "pptx_report.html"

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, ensure_ascii=False)

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(make_html_report(report))

    print("")
    print("DONE")
    print(f"HTML report: {html_path.resolve()}")
    print(f"JSON report: {json_path.resolve()}")
    print("")
    print("Top suspicious candidates:")
    print("-" * 60)

    if not report["suspicious_candidates"]:
        print("No suspicious candidates found by scoring.")
    else:
        for c in report["suspicious_candidates"][:20]:
            print(
                f"{c['filename']} | score={c['score']} | refs={c['reference_count']} | "
                f"slides={c['slides']} | size={c['size_bytes']} | dims={c['width']}x{c['height']}"
            )

    print("")
    print("Open the HTML report in your browser.")


if __name__ == "__main__":
    main()
