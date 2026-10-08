#!/usr/bin/env python3
"""Arma la extension SepararPorZonas.oxt a partir de src/Zonas.bas.

Uso:  python3 construir.py
"""
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

AQUI = Path(__file__).resolve().parent
SALIDA = AQUI / "SepararPorZonas.oxt"
VERSION = "1.0.0"
LIB = "PlanillaZonas"
MACRO_URL = f"vnd.sun.star.script:{LIB}.Zonas.SepararPorZonas?language=Basic&amp;location=application"
CONTEXTO = "com.sun.star.sheet.SpreadsheetDocument"

MANIFEST = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE manifest:manifest PUBLIC "-//OpenOffice.org//DTD Manifest 1.0//EN" "Manifest.dtd">
<manifest:manifest xmlns:manifest="http://openoffice.org/2001/manifest">
 <manifest:file-entry manifest:media-type="application/vnd.sun.star.basic-library" manifest:full-path="{LIB}/"/>
 <manifest:file-entry manifest:media-type="application/vnd.sun.star.configuration-data" manifest:full-path="Addons.xcu"/>
</manifest:manifest>
"""

DESCRIPCION = f"""<?xml version="1.0" encoding="UTF-8"?>
<description xmlns="http://openoffice.org/extensions/description/2006"
             xmlns:xlink="http://www.w3.org/1999/xlink">
 <identifier value="ar.planillazonas.separar"/>
 <version value="{VERSION}"/>
 <display-name><name lang="es">Separar planilla por zonas</name></display-name>
 <dependencies>
  <OpenOffice.org-minimal-version xmlns:d="http://openoffice.org/extensions/description/2006"
      d:name="OpenOffice.org 3.0" value="3.0"/>
 </dependencies>
</description>
"""

SCRIPT_XLB = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="{LIB}" library:readonly="true" library:passwordprotected="false">
 <library:element library:name="Zonas"/>
</library:library>
"""

DIALOG_XLB = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="{LIB}" library:readonly="true"/>
"""

ADDONS = f"""<?xml version="1.0" encoding="UTF-8"?>
<oor:component-data xmlns:oor="http://openoffice.org/2001/registry"
                    xmlns:xs="http://www.w3.org/2001/XMLSchema"
                    oor:name="Addons" oor:package="org.openoffice.Office">
 <node oor:name="AddonUI">
  <node oor:name="OfficeToolBar">
   <node oor:name="ar.planillazonas.barra" oor:op="replace">
    <node oor:name="b01" oor:op="replace">
     <prop oor:name="Context" oor:type="xs:string"><value>{CONTEXTO}</value></prop>
     <prop oor:name="URL" oor:type="xs:string"><value>{MACRO_URL}</value></prop>
     <prop oor:name="Title" oor:type="xs:string"><value>Separar por zonas</value></prop>
     <prop oor:name="Target" oor:type="xs:string"><value>_self</value></prop>
    </node>
   </node>
  </node>
  <node oor:name="AddonMenu">
   <node oor:name="ar.planillazonas.menu" oor:op="replace">
    <prop oor:name="Context" oor:type="xs:string"><value>{CONTEXTO}</value></prop>
    <prop oor:name="URL" oor:type="xs:string"><value>{MACRO_URL}</value></prop>
    <prop oor:name="Title" oor:type="xs:string"><value>Separar por zonas</value></prop>
    <prop oor:name="Target" oor:type="xs:string"><value>_self</value></prop>
   </node>
  </node>
 </node>
</oor:component-data>
"""


def modulo_xba(nombre: str, codigo: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
        f'<script:module xmlns:script="http://openoffice.org/2000/script" script:name="{nombre}" '
        f'script:language="StarBasic">{escape(codigo, {chr(34): "&quot;"})}</script:module>\n'
    )


def main() -> None:
    codigo = (AQUI / "src" / "Zonas.bas").read_text(encoding="utf-8")
    with zipfile.ZipFile(SALIDA, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("META-INF/manifest.xml", MANIFEST)
        z.writestr("description.xml", DESCRIPCION)
        z.writestr("Addons.xcu", ADDONS)
        z.writestr(f"{LIB}/script.xlb", SCRIPT_XLB)
        z.writestr(f"{LIB}/dialog.xlb", DIALOG_XLB)
        z.writestr(f"{LIB}/Zonas.xba", modulo_xba("Zonas", codigo))
    print(f"Generado {SALIDA.name}")


if __name__ == "__main__":
    main()
