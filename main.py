import difflib
import requests
from bs4 import BeautifulSoup
from docx import Document
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import sys
import webbrowser

# ------------------ Helper Functions ------------------

def get_document_url_pairs(docx_files):
    match_window = tk.Toplevel()
    match_window.title("Match DOCX Files to URLs")
    window_width = 1200
    window_height = 600
    match_window.geometry(f"{window_width}x{window_height}")
    entries = []
    canvas = tk.Canvas(match_window, width=window_width)
    scrollbar = tk.Scrollbar(match_window, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas)

    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    tk.Label(scroll_frame, text="Enter the URL that matches each DOCX file:", font=("Arial", 12, "bold")).pack(pady=10)
    for file in docx_files:
        frame = tk.Frame(scroll_frame)
        frame.pack(fill="x", padx=10, pady=5)
        tk.Label(frame, text=file, width=40, anchor="w").pack(side="left")
        url_entry = tk.Entry(frame, width=100)
        url_entry.pack(side="left", padx=5, fill="x", expand=True)
        entries.append((file, url_entry))

    matched_pairs = []

    def submit():
        for filename, entry in entries:
            url = entry.get().strip()
            if not url:
                messagebox.showerror("Missing URL", f"Please enter a URL for {filename}")
                return
            matched_pairs.append((filename, url))
        match_window.destroy()

    submit_btn = tk.Button(scroll_frame, text="Submit Matches", command=submit)
    submit_btn.pack(pady=20)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    match_window.grab_set()
    match_window.wait_window()
    return matched_pairs

# ------------------ Remaining Functions ------------------

def get_webpage_text(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")
        main = soup.find("main") or soup.find("body")
        if not main:
            return "", "Untitled Page", ""
        title = soup.title.string.strip() if soup.title else "Untitled Page"
        meta_desc_tag = soup.find("meta", attrs={"name": "description"})
        meta_description = meta_desc_tag["content"].strip() if meta_desc_tag and "content" in meta_desc_tag.attrs else ""

        # Extract clean paragraphs while flattening inline tags
        paragraphs = []
        for tag in main.find_all(["p", "li", "h1", "h2", "h3", "h4", "h5", "h6", "button"]):
            # Detect <strong> at the start of a paragraph and split it out
            strong = tag.find("strong")
            if strong and tag.contents and tag.contents[0] == strong:
                strong_text = strong.get_text(" ", strip=True)
                strong.unwrap()
                remainder_text = tag.get_text(" ", strip=True).replace(strong_text, "", 1).strip()
                if strong_text:
                    paragraphs.append(strong_text)
                if remainder_text:
                    paragraphs.append(remainder_text)
            else:
                for inline in tag.find_all(["a", "strong", "b", "em", "span"]):
                    inline.unwrap()
                text = tag.get_text(" ", strip=True)
                if text:
                    paragraphs.append(text)
        raw_text = "\n\n".join(paragraphs)
        return normalize_html(raw_text), title, meta_description
    except Exception as e:
        return f"[ERROR fetching webpage: {e}]", "Untitled Page", ""

def get_docx_text(path):
    doc = Document(path)
    return normalize_text("\n\n".join(p.text for p in doc.paragraphs if p.text.strip()))

def normalize_text(text):
    text = re.sub(r"</?h[1-6][^>]*>", "", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"\r", "", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"([a-zA-Z])\s*:\s+", r"\1: ", text)
    return text.strip()

def normalize_html(text):
    text = re.sub(r"</?a[^>]*>", "", text, flags=re.IGNORECASE)  # Strip anchor tags but preserve inner text
    text = re.sub(r"<ul.*?>", "", text)
    text = re.sub(r"</ul.*?>", "", text)
    text = re.sub(r"<li.*?>", "• ", text)
    text = re.sub(r"</li.*?>", "", text)
    text = re.sub(r"</?(strong|b)>", "", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"([a-zA-Z])\s*:\s+", r"\1: ", text)
    return text.strip()

def split_into_blocks(text):
    return [block.strip() for block in text.split("\n\n") if block.strip()]

def block_compare(draft, live, similarity_threshold=0.9):
    draft_blocks = split_into_blocks(draft)
    live_blocks = split_into_blocks(live)
    matched_live = set()
    aligned = []

    for db in draft_blocks:
        best_match = None
        best_score = 0
        for lb in live_blocks:
            if lb in matched_live:
                continue
            score = difflib.SequenceMatcher(None, db, lb).ratio()
            if score > best_score:
                best_score = score
                best_match = lb

        if best_score >= similarity_threshold:
            matched_live.add(best_match)
            aligned.append(("matched", db, best_match))
        else:
            if best_match and best_score > 0.5:
                matched_live.add(best_match)
                aligned.append(("missing", db, best_match))
            else:
                aligned.append(("missing", db, ""))

    for lb in live_blocks:
        if lb not in matched_live:
            aligned.append(("current", "", lb))

    return aligned

def format_result_as_html(docx_file, url, title, meta_desc, similarity, results):
    report = f"<h2>{docx_file} vs <a href='{url}'>{url}</a></h2>"
    report += f"<p><strong>Page Title:</strong> {title}</p>"
    report += f"<p><strong>Meta Description:</strong> {meta_desc}</p>"
    report += f"<p><strong>Similarity Score:</strong> {similarity:.2%}</p>"
    report += "<div style='display: flex; gap: 20px;'>"
    report += "<div style='width: 50%;'><h3>Draft</h3><div style='white-space: pre-wrap;'>"
    report += "".join([f"<div style='background:{'#e6ffe6' if tag == 'matched' else '#ffe6e6'}; margin-bottom: 10px;'>⬛ {draft}</div>" for tag, draft, live in results])
    report += "</div></div>"
    report += "<div style='width: 50%;'><h3>Live Webpage</h3><div style='white-space: pre-wrap;'>"
    report += "".join([f"<div style='background:{'#e6ffe6' if tag == 'matched' else '#e6f0ff'}; margin-bottom: 10px;'>⬜ {live}</div>" for tag, draft, live in results])
    report += "</div></div></div><hr>"
    return report

def format_result_as_markdown(docx_file, url, title, meta_desc, similarity, results):
    report = f"## {docx_file} vs {url}\n"
    report += f"**Page Title**: {title}\n\n"
    report += f"**Meta Description**: {meta_desc}\n\n"
    report += f"**Similarity Score**: `{similarity:.2%}`\n\n"
    if similarity > 0.95:
        report += "✅ Content is mostly identical.\n\n"
    elif similarity > 0.75:
        report += "⚠️ Content has minor differences.\n\n"
    else:
        report += "❌ Content is significantly different.\n\n"
    report += "### Differences\n"
    for tag, draft, live in results:
        if tag == "matched":
            report += f"✅ MATCHED: {draft}\n"
        elif tag == "missing":
            report += f"🟥 MISSING: {draft}\n"
            if live:
                report += f"🟩 CURRENT: {live}\n"
        elif tag == "current":
            report += f"🟩 CURRENT: {live}\n"
    report += "\n"
    return report

# ------------------ Main Comparison Logic ------------------

def run_batch_comparison():
    folder = filedialog.askdirectory(title="Select Folder Containing Draft DOCX Files")
    if not folder:
        return
    docx_files = sorted([f for f in os.listdir(folder) if f.endswith(".docx")])
    if not docx_files:
        messagebox.showerror("Error", "No .docx files found in the selected folder.")
        return
    matches = get_document_url_pairs(docx_files)
    if not matches:
        return
    total = len(matches)
    progress_bar["maximum"] = total
    progress_bar["value"] = 0
    report_md = "# Batch Comparison Report\n\n"
    summary = []
    for i, (docx_file, url) in enumerate(matches, start=1):
        full_path = os.path.join(folder, docx_file)
        try:
            draft_text = normalize_text(get_docx_text(full_path))
            live_text, title, meta_desc = get_webpage_text(url)
            live_text = normalize_text(live_text)
            if "[ERROR" in live_text:
                report_md += f"## {docx_file} vs {url}\n❌ {live_text}\n\n"
                summary.append(f"❌ {url}: Error")
                continue
            similarity = difflib.SequenceMatcher(None, draft_text, live_text).ratio()
            diff = block_compare(draft_text, live_text)
            html_report = format_result_as_html(docx_file, url, title, meta_desc, similarity, diff)
            markdown_report = format_result_as_markdown(docx_file, url, title, meta_desc, similarity, diff)

            html_file_path = os.path.join(folder, f"report_{i}_{os.path.splitext(docx_file)[0]}.html")
            with open(html_file_path, "w", encoding="utf-8") as f:
                f.write(f"<html><head><meta charset='UTF-8'><title>Comparison Report</title></head><body>{html_report}</body></html>")
            if i == total:
                webbrowser.open(f"file://{html_file_path}")

            report_md += markdown_report
            summary.append(f"{url} → Similarity: {similarity:.2%}")

        except Exception as e:
            report_md += f"## {docx_file} vs {url}\n❌ Error: {str(e)}\n\n"
            summary.append(f"❌ {url}: Error")
        progress_bar["value"] = i
        root.update_idletasks()
    md_path = os.path.join(folder, "comparison_report.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(report_md)
    text_area.delete(1.0, tk.END)
    text_area.insert(tk.END, "Reports saved.\n\n" + "\n".join(summary))
    messagebox.showinfo("Done", f"✅ Batch comparison complete.\nMarkdown saved to:\n{md_path}\nHTML reports saved alongside each docx.")
    progress_bar["value"] = 0

# ------------------ GUI Setup ------------------

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Manual Match Draft vs Webpage Comparison Tool")

    frame = tk.Frame(root)
    frame.pack(padx=10, pady=10)

    button = tk.Button(frame, text="Run Manual Match & Compare", command=run_batch_comparison)
    button.pack(pady=5)

    progress_bar = ttk.Progressbar(frame, orient="horizontal", length=600, mode="determinate")
    progress_bar.pack(pady=5)

    text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=60, height=10)
    text_area.pack(padx=10, pady=10)

    root.mainloop() 