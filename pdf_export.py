import shutil
import subprocess
import tempfile
from pathlib import Path

class PdfConversionError(RuntimeError):
    pass


def docx_bytes_to_pdf(docx_bytes):
    """Convert a completed DOCX to PDF with LibreOffice in headless mode."""
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if not executable:
        raise PdfConversionError("服务器未安装 LibreOffice，无法转换 PDF。")

    with tempfile.TemporaryDirectory(prefix="cert-pdf-") as temp_dir:
        temp_path = Path(temp_dir)
        docx_path = temp_path / "certificate.docx"
        pdf_path = temp_path / "certificate.pdf"
        profile_path = temp_path / "libreoffice-profile"
        docx_path.write_bytes(docx_bytes)

        command = [
            executable,
            "--headless",
            "--nologo",
            "--nodefault",
            "--nofirststartwizard",
            f"-env:UserInstallation={profile_path.resolve().as_uri()}",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "--outdir",
            str(temp_path),
            str(docx_path),
        ]
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=120,
                check=False,
            )
        except (OSError, subprocess.TimeoutExpired) as error:
            raise PdfConversionError(f"PDF 转换失败：{error}") from error

        if result.returncode != 0 or not pdf_path.exists():
            details = (result.stderr or result.stdout or "未知错误").strip()
            raise PdfConversionError(f"PDF 转换失败：{details}")

        return pdf_path.read_bytes()
