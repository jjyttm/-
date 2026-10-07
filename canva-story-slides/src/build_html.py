import base64, re, subprocess, sys, pathlib
src = pathlib.Path("slides.src.html").read_text()
chars = set(re.sub(r"<[^>]+>", " ", src)) | set(src)  # include text in attributes/JS strings
chars |= set(chr(c) for c in range(0x20, 0x7F)) | set("·→←✓⏱▶“”‘’")
pathlib.Path("chars.txt").write_text("".join(sorted(chars)))
subprocess.run(["pyftsubset", sys.argv[2] if len(sys.argv) > 2 else "fontpkg/package/dist/web/variable/woff2/PretendardVariable.woff2",
  "--text-file=chars.txt", "--flavor=woff2", "--layout-features=*", "--output-file=PretendardSub.woff2"], check=True)
b64 = base64.b64encode(pathlib.Path("PretendardSub.woff2").read_bytes()).decode()
face = ('@font-face{font-family:"PretendardSub";font-weight:45 920;font-style:normal;font-display:block;'
        f'src:url(data:font/woff2;base64,{b64}) format("woff2")}}\n'
        '@font-face{font-family:"Pretendard";font-weight:45 920;src:local("Pretendard Variable"),local("Pretendard")}')
page = src.replace("/*FONT_FACE*/", face)
pathlib.Path("canva-story.html").write_text(page)
title = re.search(r"<title>.*?</title>", page).group(0)
meta = re.search(r'<meta name="description"[^>]*>', page).group(0)
body = page.replace(title, "", 1).replace(meta, "", 1)
standalone = ('<!doctype html>\n<html lang="ko">\n<head>\n<meta charset="utf-8">\n'
  '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
  f'{title}\n{meta}\n</head>\n<body>\n{body}\n</body>\n</html>\n')
out = pathlib.Path(sys.argv[1]); out.parent.mkdir(parents=True, exist_ok=True); out.write_text(standalone)
print(len(b64)//1024, "KB font;", out)
