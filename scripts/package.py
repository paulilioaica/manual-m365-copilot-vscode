from pathlib import Path
import json
import xml.etree.ElementTree as XML
import zipfile

root = Path(__file__).resolve().parent.parent
output = root / "dist"
output.mkdir(exist_ok=True)
manifest = json.loads((root / "package.json").read_text())
target = output / f"{manifest['name']}-{manifest['version']}.vsix"
vsix = XML.Element("PackageManifest", Version="2.0.0",
                   xmlns="http://schemas.microsoft.com/developer/vsx-schema/2011")
meta = XML.SubElement(vsix, "Metadata")
XML.SubElement(meta, "Identity", Id=manifest["name"], Publisher=manifest["publisher"],
               Version=manifest["version"], Language="en-US")
XML.SubElement(meta, "DisplayName").text = manifest["displayName"]
XML.SubElement(meta, "Description").text = manifest["description"]
XML.SubElement(meta, "Categories").text = "Other"
props = XML.SubElement(meta, "Properties")
XML.SubElement(props, "Property", Id="Microsoft.VisualStudio.Code.Engine", Value=manifest["engines"]["vscode"])
XML.SubElement(meta, "License").text = "extension/LICENSE"
targets = XML.SubElement(vsix, "Installation")
XML.SubElement(targets, "InstallationTarget", Id="Microsoft.VisualStudio.Code")
XML.SubElement(vsix, "Dependencies")
assets = XML.SubElement(vsix, "Assets")
for kind, file in [
    ("Microsoft.VisualStudio.Code.Manifest", "package.json"),
    ("Microsoft.VisualStudio.Services.Content.Details", "README.md"),
    ("Microsoft.VisualStudio.Services.Content.License", "LICENSE"),
]:
    XML.SubElement(assets, "Asset", Type=kind, Path=f"extension/{file}", Addressable="true")

content = XML.Element("Types", xmlns="http://schemas.openxmlformats.org/package/2006/content-types")
for ext, mime in {"vsixmanifest":"text/xml", "json":"application/json",
                  "js":"application/javascript", "html":"text/html", "css":"text/css", "md":"text/markdown"}.items():
    XML.SubElement(content, "Default", Extension=ext, ContentType=mime)
for file in ["LICENSE", "NOTICE"]:
    XML.SubElement(content, "Override", PartName=f"/extension/{file}", ContentType="text/plain")

with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as archive:
    archive.writestr("extension.vsixmanifest", XML.tostring(vsix, encoding="utf-8", xml_declaration=True))
    archive.writestr("[Content_Types].xml", XML.tostring(content, encoding="utf-8", xml_declaration=True))
    # Explicit allowlist to keep editor settings, credentials, and scratch files out.
    for name in ("LICENSE", "NOTICE", "README.md", "package.json", "extension.js",
                 "chat.js", "continuation.js", "prompt.js", "socket.js"):
        archive.write(root / name, f"extension/{name}")
with zipfile.ZipFile(target) as archive:
    assert archive.testzip() is None
    assert "extension/extension.js" in archive.namelist()
    assert not any("node_modules" in name for name in archive.namelist())
print(target)
