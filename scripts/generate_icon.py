from pathlib import Path
import struct


def _row_bgra(width: int, y: int) -> bytes:
    row = bytearray()
    for x in range(width):
        # Gradiente azul
        b = 190 + ((x + y) % 50)
        g = 110 + ((x * 2 + y) % 90)
        r = 30 + ((x + y * 3) % 60)

        # Moldura clara
        if x in (0, width - 1) or y in (0, width - 1):
            b, g, r = 240, 240, 240

        # Letra N simplificada
        if (3 <= x <= 5 and 3 <= y <= 12) or (10 <= x <= 12 and 3 <= y <= 12) or (x - y in (0, 1) and 3 <= y <= 12):
            b, g, r = 255, 255, 255

        row += bytes([b, g, r, 255])
    return bytes(row)


def make_simple_ico(path: Path, size: int = 16) -> None:
    width = size
    height = size

    # BITMAPINFOHEADER (40 bytes)
    # ICO DIB height = real_height * 2 (inclui máscara AND)
    bih = struct.pack(
        "<IIIHHIIIIII",
        40,
        width,
        height * 2,
        1,
        32,
        0,
        width * height * 4,
        0,
        0,
        0,
        0,
    )

    # Pixel data (bottom-up)
    pixels = b"".join(_row_bgra(width, y) for y in range(height - 1, -1, -1))

    # AND mask (1bpp), alinhada em 32 bits por linha
    mask_row_bytes = ((width + 31) // 32) * 4
    and_mask = b"\x00" * (mask_row_bytes * height)

    dib = bih + pixels + and_mask

    # ICONDIR
    icondir = struct.pack("<HHH", 0, 1, 1)

    # ICONDIRENTRY
    b_width = width if width < 256 else 0
    b_height = height if height < 256 else 0
    entry = struct.pack(
        "<BBBBHHII",
        b_width,
        b_height,
        0,
        0,
        1,
        32,
        len(dib),
        6 + 16,
    )

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(icondir + entry + dib)


if __name__ == "__main__":
    out = Path("assets") / "busca_nfse.ico"
    make_simple_ico(out, size=16)
    print(f"Ícone gerado em: {out.resolve()}")
