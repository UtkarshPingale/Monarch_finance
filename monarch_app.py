import os, socket, socketserver, webbrowser, traceback, secrets, html
import threading, re
_req_ctx = threading.local()
import csv, io, math, hashlib
from http.server   import SimpleHTTPRequestHandler
from urllib.parse  import urlparse, parse_qs, quote as _urlq
from datetime      import date, datetime, timedelta
import calendar
import tempfile, traceback
from sqlalchemy              import create_engine, text, event
from reportlab.platypus      import SimpleDocTemplate, Paragraph, Table, Spacer
from reportlab.lib.styles    import getSampleStyleSheet
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen        import canvas as rl_canvas
from reportlab.lib.colors    import Color, grey
from openpyxl                import Workbook


# ═══════════════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════════════

DB_USER = "postgres"
DB_PASS = "Monarch%40123"
DB_HOST = "192.168.1.149"  
DB_PORT = "5432"
DB_NAME = "postgres"

SCHEMA_NAME = "monarch_db"           

ENGINE = create_engine(
    f"postgresql+psycopg2://{DB_USER}:{DB_PASS}@{DB_HOST}:{DB_PORT}/{DB_NAME}",
    connect_args={"options": f"-csearch_path={SCHEMA_NAME}"},
    future=True,
    pool_pre_ping=True,
)

# ── Live "data changed" version counter ─────────────────────────────────────
# Every INSERT/UPDATE/DELETE that goes through ENGINE (i.e. every write the
# app makes, anywhere) bumps this counter. Pages poll /api/data_version and
# auto-refresh themselves the moment it changes, so any change made by any
# user shows up everywhere in real time without anyone needing to hit F5.
_DATA_VERSION      = 0
_DATA_VERSION_LOCK = threading.Lock()

def _bump_data_version(conn, cursor, statement, parameters, context, executemany):
    stmt = statement.lstrip().lower()
    if stmt.startswith(("insert", "update", "delete")):
        global _DATA_VERSION
        with _DATA_VERSION_LOCK:
            _DATA_VERSION += 1

event.listen(ENGINE, "after_cursor_execute", _bump_data_version)

def get_data_version() -> int:
    with _DATA_VERSION_LOCK:
        return _DATA_VERSION


SESSION_COOKIE   = "monarch_session"
SESSION_LIFETIME = 60 * 60 * 8
SESSIONS: dict   = {}
TOKEN            = secrets.token_hex(8)

C = {
    "primary":      "#000e24",
    "secondary":    "#775a19",
    "bg":           "#f7f9fb",
    "surface_low":  "#f2f4f6",
    "surface_high": "#e6e8ea",
    "outline":      "#c4c6d0",
    "on_surface":   "#191c1e",
    "on_var":       "#43474e",
    "error":        "#ba1a1a",
    "sec_cont":     "#fed488",
    "dark":         "#1a2540",
    "mid":          "#2c3e6b",
}

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
LOGO_URL  = "/logo.jpg"

# Logo is embedded directly in this file as base64 (instead of a mapped network
# drive path like Z:\... or a separate image file on disk). That old approach
# broke the logo for anyone running the server on a different machine/session,
# since a drive letter or relative file path can go missing depending on who
# launches the app and from where. Embedding the bytes here means the logo
# always travels with the code -- nothing external to lose.
LOGO_B64 = (
    "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/"
    "2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCADIAMgDASIAAhEBAxEB/8QA"
    "HwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkK"
    "FhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXG"
    "x8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAEC"
    "AxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOE"
    "hYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD856KKK/YD"
    "IKKKKACiiigAooooAKKKKACiinJG0rhUUux6KoyTQA2ivSfh/wDAH4i/EG1ubrw74Q1HVreJQXkh8tR16Dey7j7DJ9q47W/D2p6Zqb2l3ZS29zGdjxOMMrAk"
    "EH0IIIxWMa1OcnCMldb6js7XMeinzwvbTSRSLskRirLkHB/CmVsIKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAFRdzqvqQOma99/Z"
    "3/Zr1L4ifEC2OoFU8LWLsdRvkQlCoBBjjboXJKgemSf4TXl/gnwb/wAJNfW0IjuTJJJGsZhI+d2OI0HHUvxn0Br9U/CPhyHwH4Vs/D9sYswgNcmFcRyTlf3j"
    "L6gH5VJ5wCe9fG8S53/ZOG/d/HLRf5ndg8M8RO3RGjbWdppul2Wm2FrHZaZZxCK2s4hhYFHYHux6lupryT9o/wCAdn8YvD9zqulLFZeN7VNyzYOL+JVJZWVR"
    "80oGSD1boecV66gAzgY9fepI5GiYMrFWU5BHUH1FfgmAzrE4PGfWlK7b18z6mrhIVKXIkfkLrPhC+0WEy3qNbOZWTyZY2DYHBYkjpnj1zWDX3x+2h8ILC58L"
    "zeN9NtTDdNPt1aRSzfO7ZEmCcASEgHjAZT618HXlutvLsUkj1Nf0vlmPp5lho4ins/wPjatKVGThLoQUUUV6xkFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFF"
    "FFABRSopd1VRkkgAe+a6/XPAviHStNs7m90e7tbXUFD2k8se2Occ8o3Q/dPftUuSi0m9wOPoqe5sbi0YrNE0ZBwQe1bXhjw7e67eafBaQNPJJc7ERGGWbjAA"
    "zRKSirgfV/7DXwuOoC48c6pbx/ZNOaO30yEcFroLnzj/ANc1JYf7TZr7CVQi7VGB1xWB8OPDr+D/AIfeHNCmsnsLu0tl+2QPt3C5bmbcQSCc459APSuiPSv5"
    "h4mzSWZY+bT92LsvkfZ5fQ9lST6sbTl5ptKpxXyJ6gzUNK07XNOu9L1e1S80vUIXtruCQZV42GBxnscMPdRX5d/G74by/DLxnrnh+5WOSTTpjGt1GOJQSpjI"
    "9ih3V+pDHOfevmP9uPwM9/4R0rxfaQrmxcWGpSEDO1iPJbP1yvcn6Cv1LgfNHQxLwU3pPb1R87mlBW9oj4Goq09rLO00iJlEOWO4DH51IuiXzwecluZYuPni"
    "YOOfoTX722lufNFGitzVPA+v6LaWVzqGkXVlBegtbPcJs80AkEqDzjIPPSsWSNom2uNp9KFJS1TAbRRRTAKKKKACiiigAooooAKKKKACiiigB0bbJFbrtYNx"
    "7Gv12tl0ib4daPp+px2l54ebT7VmstTiWS25VWX5Dwp3HqMc1+RUGPPiz03rn6ZFfqpf+HT4w+EelaXp86ATWFsyvMdwIXBXp9MV+Zcc1KtLDUalJtNS6H0W"
    "R0KOJxXsq7tF9Tzb42fspeD9c0K91zwuJdDurSCe9l09ZfNs5EjiLnygRlD8p4JI9MV4V+yR4bs/Ffxl0G2awA/s55dQuRjYGijTcF+u5Qc19YTaNqHgv4B+"
    "J7W8k33ltpN00r79yhHIQjPX7rmvB/2G0hPjnxJd7pJJ20qVI5ZMZwJI1A6ehYVyZTmeJrZLiK1eV3G6X3BmmEo4XHOjRd433PsYMXw7ne7MXZvUnk/qaaww"
    "KQKFAAJxjikzX4XNtybPooLlikFFFFI0CuX+Kngz/hYnw18QeHzLDC08AuFknVmUNC4l4A53EKyj/e5rqKuaP/yFLQcfNKqHPTDHaf516eWVnh8ZTqx3TRx4"
    "qHPSaPzE+FPwZn+JPxJbQ0LWkRuTHPdeWWFqEyXyO5IBAHrX3z4N+Ffw/wDhTYRrpel228FZBfa3smkEhxzGmNiHPsT7mvkn4F6g+j/tQWNnJI0Vla6vcR3O"
    "xiDI290G7+8csv619TfEX4baj4w8V6Pdw3MMFtbRR+eJC2Syvu4A68cV+q8ZZljKHs6dGdoyXQ4cgwOFxtWSxMuVK5yP7aot5/hfpF3eo13dtqctvHMSMohg"
    "Q4B7AMc4HHWvzv1S2ZJS5YHpX6Fftpzxf8Ko0USHav8AbcgUA8g/Z1IBJ/3TxX5+6x0b8K+y4TlOWV05T3d/zPAx0YxxEow2RkUUUV9mcAUUUUAFFFFABRRR"
    "QAUUUUAKAW6An6CtTStHa/jPGJC2FQqSxGOoHU0mhvHGL5pDs/cYWTJG1iwwePz/AAr9Of2fZLL/AIVJ4aubIWwn8krKVtk82N927YZCuSOd3BwN1fP5zm0c"
    "ow/t5RujooUXXnyI/OWH4f6thJY9OuT0ZSLSTtyO1fSfw5/a08W/D3w3Y6Drejx6yUQx20N7Bj7KozhQwZXxyeO1fbx17UuQL+cD0VziqN4w1S3eDULe21CG"
    "T76XVvHIG+uVr8yxHGuCxsfZ4rD3j6nsLLK1PWEtT5A+Jn7Ww8WfDbWNDh8PSaRqd3EqSywXXmQPACWlXDAEZwvGT0rS/Y8sEi8Wa5LBCkcK6OrqsY5/1gLN"
    "gevc17z4l+CXw78WwSx6l4L0uEzAbrjSlNpMCDnIK/Ln6g15tq37L+o+CtRXxH8Itdlg162GYtNvGWC4l9V8/OyUf7DAZ5rvw+bZRi8JUwOGfsue++12c8sN"
    "iKE1Vmr2PeM4OM8CnkYrx/wp8fdS08/2T8WNAm8M6tCGSXVoYAkMj5AxLCvAPOAUPPpXscqOkdu+3fFLGHjnU7lkQ9CCOxGD+h5r8uzHJsVl0r1FePRrZn0W"
    "HxUMSvd0fYjooAJ6VLbxNPcpAsbO7gnHAXA6lmPAA6knivDjGU5KMVds65TUVeRFj2qezzHNHcsRFbQyq8k8jrHHGFIZiWYgcAE14j8R/wBqLQvC9xNovhRY"
    "vFWvwb0nmMcht7R1ODtTA80jHUsF789K4fw98Pvi58erzT/EHi/W30jwpODsguH8osmCA0duBkDPoq55Ga+/wHDM6cFiswqKnHz3PGq5hzN06Ubs8OtPEF5p"
    "nxjm8U6VZ213GutyXXkoeJyHaTg9QXTABr6E8SftplluBoPgeKGXzG/e6tdNNsG45V0QKoYHtuI6cV6T4c/Z2+GvhCfzrTQZdTvVAButQnYAsP4hGhGPxY5r"
    "uoNO02wfdZaPpVlkYKwWEaj9Qfzr6PMuI8mqygqlL2nLojzqODxKu4u1z86fi58SfFHxEukn1jUru8hjlea2tMloLXe3KxKBgLjA78AV5PrNpPhz5E2Byf3T"
    "dPyr9eFu2jVRGqJgYASGMAf+O1bsru6ubuCBHBeaRYsNFGQQxwQQV5HNdOH45wsOWlToWQ55ZUs5ykfjE8bR/eVl/wB5SKbXovxwtE0/xrq9vDcNcQLql5sJ"
    "6ACTaAPbAH8q86r9bpT9pBT7nhvR2CiiitQCiiigAooooAKKKKALNnlxLGDjeAM/jX6HfsefEmw8T/D+68LpZyw6josi3U0zyKVkhlCxjaByMMmefWvzy0+G"
    "eaZhAoZgMnP1r6P/AGTPEqeCPipoEuoTR2w1id9IuASV/dSKNjE4Iwr7TjrXzPEOCWPwFSk97XXqjpw1T2dWMj78opZUa3leJxtkjYow9GBwf1pK/lqUXBuL"
    "6H3ifMroazUgCtyRkjpQeTSqvFSpOOw7JqzK3iXRtM8c+H20TxFa/wBoWJJ8mUn99bN1DI5ycA4JU8emK+c38X+NP2XPEdrpHia+PivwhdL9oa4hi2pcKWCB"
    "oWP+qkDFSR3x8w719L4xn6YqlrvhzTPHPh258Na7b/atJvWXcNoLwvn5XjODtOcZx1GRX2uR517L/ZMXHnpS0s915o8nFYNfxaOkkT6bqFhreg2ut6bfRXOh"
    "zxs4vWOxIQPv+bn7m3HOfqM5FfMfxm+P2teNdfvfBXw3tF1DT44Y2kbY/nak2SCpGclCSNsffnOa+f8AxdY+JPB+o6/oKXF3BameXTvs8d/I3yxydXjzghcZ"
    "AOc5r7K/Z8+Alj8J9Pt/EUr3k3ijULXKrcSKws4JArYxtDLM3OeflU4HU19z/ZmW8Nxlj5vnb+FM8iVetjWqXYf8AfgbZfC7w6mqa7Y21/4y1FRcXVtOgkh0"
    "5ic+XtPDyj+Js4B6etesSSSTzPLI7SO7Fmd2LMxPck0wKVJII9Mj0py8V+V5pmlfM6zq1ZfLsfRYfDRw8UluRFfmNG2nMc0leK9TtE2ioNR11PCOm32tzxNL"
    "HpsD3BVWxhwCE5/3itWMEnA4J6V4N+2B45vPCXgfTdHtWEcuruzzEqTmJRhRweMks/Oc7QR0r6Lh/Af2hmFOn0vd/I87HVVTotdz4S8ZzS6xf3F0824TTzXX"
    "zDnLtk/rXJV0jMbhJYXA2wjCkDk5PeufuIxFKVXOMd6/qamuVcqPiGR0UUVqAUUUUAFABJwASfaitLQrOa61CHylzkkZyBjAySfQD1pN2VwHWmnSTRx7bSSR"
    "iu44iY8A8np096+j/hv+xj4x8daTPrF2mkeH7FpPLT+0Y9u/pzGoBboc5AwfXmvob9n/AOAVl8P9C0vXvEFmb3xTc2yv9luMmGwQ5Kq65w745/ujIyCa9ulu"
    "JbiXzJnMsrD5nckkewPavynPeM1g6jw+EV5Ld9D2cJl7re9N2R4B4M/Yp+HvhZU/tS51LxNcjBfBWyhZs9PlBkKdsFhXouq/BDwnq2gPpmieG9K0TUoA8unX"
    "UUOSs+0gq7sSXV1G05PB2nqK7Yg5pQp5IbaccH8a/NZ8T5lUrqrUqOye3Q9pYCjGDilqcN8JvG0/jjwdEt9DNB4j0nZZarFNIWZpOdkwJ5IYDB9GX3rt433V"
    "y/jNZvCOov4807TvtgmxbeJLZJQplhP3Zxk8Ddg4H3WweAxrp7a+stUsLbUdMuxe6TdIJLa4CFCw6FWU8qwPBU8g0s5wsKqWYYVfu57+T6ojB1ZRboz3Q9qS"
    "lIzjHPuKTBr5M9dDuoxnBNcf8VfihB8JPBd9rpiFxqBcWmm22D++uWU857BF3P8AUCuwDBcZ6DrXyL+2d4rk1nxxoHgkzzxQW0azMkIAMlxOPlIboCFZEyxG"
    "Oa+w4Wy9Y/MY8/wR1fyPJzGu6VOy6lb9l/4f3PxP1mPX/EVxNeWWiXI1C5jugXbUJmbEcTk/eUEBiewX3r7FmnkmmlkeRpGkYsWJ+971yfwk8DTfDX4XeHvD"
    "13AlvqcUAnv1V9+Z3A4J/wBlAi+nXHBrqWJB4rbinNHjsbKnH4IaIMvwyo01N7sXPFNyfWm5bNOr4o9cKKTOKdEpmlSNBvkdgqoOpJ7fjQld2Qm7K7I5Z4bS"
    "Ca4ncx29vG88rhdxWNRlsDv6D3IFcZ4YstG+MOn3fivxL4dstRiu52ttMsNSiWeKC2jbG8IeA2fkDD/bIzmm65qb/EvV7vwZpCtJ4Ytg6eJ9aVtnlyKxC2cI"
    "OCzblPzDKk8/dXnt4hDBBBa2ltHZ2FrGsFtbRDCxRKMIg+gr7a/9iYO17Vqn3xj/AME8S316tr8CPNfEn7MPwt8Q2kyQ+F4vD9yw2C60noT/AHjHJkH8CK+e"
    "vif+wzr2lSC48E3Vj4n0yOISTQyxrHfggZfEePmUdgrE9eK+0SM05CUZXB+ZTkc4x9DVYDi/McHJKUuePmaVctozXuqzPx71bQ72w1C+hayuUW3meNibaRAm"
    "CeuRleOx5rO2nBODgdTiv15+JfgzTvin4cfTNXkltZn3eVqUYy8bFdp8wD/XJjqDyOx7V+anxy+HN78PPHeqaFM6zTWrIheFcRyJgFHB6YKkEZ/pX7Xkmf0M"
    "5heCtJbo+Zr4eVCVpHmFFSLA7lgByOvNFfVHKR1o6ZqDW8LW8a7WlkBaQHB2fxJ9Dxn6VnU6OQxOGXGR60mroD9APgh+17oOvaTZ6T48uV0m7tf3EWuMm4zk"
    "KAqzIg9P41GfUd6+kIpIbizivbW5gvrCY4hvLWQSQy/7rjj8OvtX48x6nIk3mNHHKwGBvBwv0wRXovwt/aF8a/DLUIl0TVZLexdh51isYkimAyRmNsqfyr81"
    "zrg2hj262GfJN/cz1cNj50PdeqP1E7cjB9D1orwDwB+2p4M1/Tol8VxTeH9S482bT7d5rUk99h+Zf+AlsY6V7n4Z13R/Gtsbnw3rWn+IIOhWwnDSqfQxnDg9"
    "OMZr8hx3D2YYCTU6ba7rY+hp46jU1vZl4Nt+8kcqFWVo5F3K6kYKMO6kE5FeYtZn4APqWrWMF9q/w91GZGm0uFg82kzMwXcC3VckBWHUHaxyAT6fMvkHbKHh"
    "YHBEqlSD+NKly0T7yFdSCrLIoZWU9QQeCD6Gs8tzCpgb0a6vSlvF/p5ixFCNZKUHaXcr2N3baxp0Wo6ddrf6dJlUnRSuGHVHU8o47qf1HNPDbhkCuH1r4f61"
    "4Yv7/wAQ/C+6+z6ncBmu/Dt5IJLW6CjI8tDgPnkBScjsa53wZ+074a12ddJ8WQS+AvE6Hy5LW+RvspYdcMRujyc8MCB616dXIFioPFZY+eHb7S+RjSxrpNU6"
    "6sz1tU3OgKGQE8opwSB1A/DNfF3gizX41ftUeZqLeZZ3F7cX8rBf3axxguVx6YjUfhmvsDWdYSx8G6nrdnfQSW6WEr21/aSpLA0roViIcHGCzLwcZr5y/Y48"
    "BtDeeI/EFzG8clmTZKhkyVd8BiTjkMPM49q+k4dhLLsuxWKmrStZX/rucGOqRr1oRjqj6emuftF1LIerkvj6njHtjApG60nTgAfWnEV+V1HKpJye59LFJRVh"
    "tITgVPa2c947LBG0zDnbGMkD1PtXK+O/iX4S+Gdt5niDV0muGcRxabpGLu7kbP8AdQkL9SfoD0ruwuX4rGTUKMG2zGpiaVNXlI6UQyTNtiRpSBkgcfXrxgep"
    "rhf+Fgnxx4gu/C3gS7Zbmy2vq3iaSBvstrFuw8ERIBLsAR5g9goO7cG6dfeLfidCouLU+CvA84WVo5VLXuopnBQMcH3wQF55zxXXaF4f0XwlpY0vw/pyaZpg"
    "kMvkg7nkkIwZJH/jc4HPQdAAK+pjhsLkkXOt79foui9fPyPKdSpjZcsdI/mM0rQNK8MWH9m6LZfY7ASvMwc5kmkbrJIe7EAewAwKvx047SQMnJ9KW+WPRNOl"
    "1HU54dK06MgNd38qwRAk4HzNjPPpXy1T61mNZzcXKTPVXssPFR2Qo5IHc9KWGN7iTy4UaabOBGgyT+VeVeKf2qPhp4OinV9QuvEN9HGHW00sGKN8khQZXA9M"
    "5UHg9a+Wfi7+1r418bW1zYWfleG9EnVf9CsQVc8Hl5Actn0OQOwr6zLOD8fjJJ11yR8/8jza+ZU6ekNWfXnxI+PnhL4VQstzOviTVgxH9k6fOFjjK9RNNgqC"
    "CMbVyc55r4A+NnxEuPiT4tufE87TI06ojwOy84yB90BQFBAAx0FcLda7eXW57iT7TIF2h5iWIAGAOtZM+ozXMXlvt2nB4FftOUZFhsojaitXuz5yviJ15XkN"
    "juRG7ttJ3e/SioKK+mscoUVJBbyXLlIl3MFLEZ7DrV9PDepSW5uEtWkhzjehBBPt60NpbgZlPhlkglV4naOQdGQ4Iq8vh+/cDEGMjcAzAHGcD8/SnQ6HfR3Q"
    "XyC7KxBEZDHv6dRxS5l3AXStVv7O4aSG8uIX2/eRyDXSaB4t1XSZhPb6ndW8wk8wSRzsjBuOcg5z71ivp10n34HQerjH86lfRrmOIO4hUFN4BmXJH09faspK"
    "E1Zge6+Hv2w/iPoEEaprlxe21kog2ag4vA49SshP59q7NP8AgoD4iW4jN14P0G8tIyfMALRyP6fPG2Bzj+GvkyTwrqUsU10tpmFMlmLKCo9/Sqlhp8s0wKBG"
    "Kk5XeN3TqB1Iry6uU4DEa1KUX8jZVakdmz7fsv29YrmLfP4BtIwB/wAumrSmQntwwwB71wXxx/aA0v46RWFq3hiy0zW7YAS6wLh7iR0AO2IO2AVGc5bJB9q+"
    "a42ktixxtOdpDDB9atCCaeHzpFUJu2bmIAzjNY4fJ8FhJqpQhyvyCdWdT43ctXfinWtMnuNKfUbx7eOVI5I47hmicoRtOAdrYwMH2Fe2/Bj9qq++Emlanpt9"
    "ocOvebKLhnkvGhIxwGbaDvOGwMdMmvCksLmRkSK3kkMkYkQRoSCpHWqeoaRcwWsbMiFZV3oVcEY9/Q/WvQr4WhiqbpVY3T3M02ndH143/BQqaTK2/wAPdNZh"
    "3l1Sfj/x0VUX/goF4hRsJ4Y8PWqN0URvcSD6F2x+Yr45+xS7FbZkNnGOen8qv2dnNFaK7xmON5DGGbhd1eZHIMrhtRR0PFVmrczPf/F37XnjnxTJcwRapc2V"
    "tIzsIbecxIisCCAqnA44wOK4r4b/ABHtPCfiC21a88PWHiiO0DynTdRysMp2EBmPqv3h6kCuAudLuLZ2V0VsMF+RwwJIyMY60j6XdIMvbuoyBlhxknAr044T"
    "D04ezhFJPsc7lJ6s+zpf2+t9raXN18P4riabI3/2054GBgDG4Y9DWZrP7e9/aru0/wAH6TalhgR3bPPt987/AOlfI0WkXMsMkyxDy0YozEgYYdqztV0u4tLl"
    "FliEZeNXAyOhzjpXlLh/LHPmdJNm/wBYqtW5j6bvP25fiTe20kFnJYaZG4Y77WFIn6fwsvzfma8s8UfHnxf41s3ttd1S/wBQhfBJurl5Twcj7x9a4G3sZHhj"
    "x5ZHAJ3j5f8Ae9Ks22i3t2ZxDAXaCLzpFBGVTOM4/LjrzXpUsDhMO/3dNL5GLnKW7JTrgZo2KMWjOUJ5Kn2NQalqk2qXBlllkkO0L+8bJ4ot9Jubq5SBFTzX"
    "4AaRV57A5PBODgHrT5tCvreV4pYPKkQ7WR2AKn0Ndy5UQUdqkcgc1BdxIIDsQZyOgrXfQL6OAzND+7C7t28dM4/nVFIncgKjMT2AOapSQGQQR1BH1FFXdVhk"
    "t5ESVDG4zlW6iitFqBWtphBMrldwDAkA4OAeee1fot+yd4v8H/Ev4dTLbeDdJstW0SRYbr7VBHcSXatz9pJ2juwUr6gY4OK/OSvtL/gm/pM2ra34ojWaQKLV"
    "kSCJQfNZ9gKsT06KwI6FD618zxFS58vqTUmnHVNHRh2lUVz234//AAhsfiV8M9StdI0jTNN1yzIvreaxsRC8pU4eP92Bu3IeM/dK5qh+yp8Oj4Q+Elpc634c"
    "0+TVr+4ml+0ahYJLMYhtVBlwfl4OMYzXtVjfPp91HcK5UxvuJ6HA4/lxWnrM7XNtbdVWCIIisc7UKKVX6gYFfidPPMU8snhXJ3TvfrY+glhIRrxdtH0PE/2j"
    "dK8K2fwg1i9vPCem3F9K621i1nZpDJBOylvNDLjoFxjvurhf2bv2Y9P8N6LB4o8Z2VvrGqahEJ7TS7+3YJaDOVklG7DlgDheAAQT6V6F+0RaxXngfQI50EkH"
    "/CQweaD3Ty2DfpXqmrYXUrlUJMUchjjB/ur8q/oor1Kmc4nCZLS5ZPmqN3fp2MaeFjPEyT2R57Yaj4IufijP4UHhfQW1BtL88n+zIdpYMXKbCuNoU4z97PU4"
    "4rzT45/sveGdX0mXXvC+jw6bf2X+kXOlWaHy5U3AGSPn5AozuTkc5GMGodAcv+2ndOfvixnXPsIyBX0TAsTyxpcLut3BjlXpuRhhh+IJFdeKzHEZTisK6c21"
    "KMW03vczpUIVqdRvdbHlP7PFr4L8S/C3SpLfwro8Gq2EAg1C0vrKO5lV+czbnBJVm49iMdK7XxR8LfC/jvR59Hl8P6bZzzxt9muNOsYbWVJgpwhYAAq3K8+o"
    "NeX/ALFaadZ6x8QI7m2KLb37xSlRlltwkzgKCfoSO5A9K93v7b7FchN5kQorRyjpIjDKsPw/UVy55icTl2Ye2ozk43Wl9LvU0wlOnWpOMtzzf4Q+GvC8fg/T"
    "LNvDOlwavpcH2e+mvrFZsSBfnny2flPTb/CRx1r4h+OXiPQvEnxY1uPQ9FWx8NwPFsghOG3DaryKh/vMc7BwBX6Fa54YXStF8ceNNMZ5LybTS99EWG2J40cI"
    "6L6EEZ9x1r8zJtMXWPGto7JJNe3kvlIIzgmVhtTHp8xHNfeZBVjialXFKTaduu2l2eRWg6b5Zbn3b+y/8P8ASvCfwpsZ5NG0+9u9Sk+0TXF/ZRzjbtxGE3A4"
    "+6+enP0rO/a68C2fiD4SHUrHSNMsL/S7ncslnbJAxjdCCMKADgrnJ5/OvUvDs9v4eisfAsQTbpOlxyIAoWQFTsk3Y75+b3OT3q14r0SHxN4S17TLiOSaO4sp"
    "QI4ThndB5kYB7fMgz7Zr4F5xVhnkanM+Vy76W2O6iqOIwklT+KL1Pi/9i19Mv/ipd6dqumWWqr/Z8v2STUrZZcSZjfOxgRkYx619tXml6LqdtLay+FtAkjnU"
    "oVXS4UJzx1Cgj6ivhr9mfdZ/tAafArZEc86bv7wZecnvX3VF8zIFbY5wFb0PY16nF2IrUMwoezm0pWvqVgaUakJOS2MOy8F+BdQ8LaekPhTQLvSmWWMwGzz5"
    "cgOyRGfO5iMD5ic854NfHH7V/wCznpvgfxFour+GCLTw/rB+yR2QVpWtp0GXG4nLAhgQOoHGTX2V8OfD13pHgHQ9HujHLqO+eaQQkHPmy7kzjjcUCk+55rmP"
    "iJa2ms/Gf4f+CdReIJZXrX9wkR3M8oj3bBx8uwRJu7HdgVGVZjiqGa1Yxm5Uo3b67IqvSh7CEras5j4K/sy+HPh7odlPrOn23iLX5o45lW/g3JZEqCV2n7zE"
    "8kNkJ0wTk1tzeJvhfH8WdR8Jz6BoovJnEEt/JaRxRtcR4CgKiggPI23cpB7sSOnqAnkuJRPKcySNvbJzyTk/rX55/Gi8ktvjX4llN+bRk1CW4Fw672VlBIVf"
    "TJatsixFbP6+IlXqNW2s7WM8VQWGhC3U+if2kPAngjwlb+E7+48L6Xp1rY6qXuLaxt/KeZd8e5JH+84ODy5YjJAxk17jf+BPB8d40cfhLw9JEAvluNKhG5CA"
    "VPT0NfGPxY+P3/C1fB/g7w6+ntHdw7I9Wu5ArRzyYUAwEHKfdyc/3vavtwIEt7EAk4s7YcnP/LFa5c/+u5dgKHPUaneSevS+heEjTrVZaaHOvD4Jh8cWfhtv"
    "COmDV20dr6K9FhF9nWMuf3Bj6EkIW3446Vp3WieHba1ml/4RTw8QiMxH9kW/IAJI+76VRvPAt3efEq28XRyRQ2Vlocdm4kBdppGDkBQB8uA33mI9q0dZYro2"
    "odm+zS8f8ANeJjsZVhUw0adR+8o31e7OmjRjNVG1tc/ML4/6hoOrfEa/vvDmjDw/plyfMTTUl8xIT904OO5UnHQZwKK53xzEovI5eS7Fgcn3z/U0V/Q1FctO"
    "KR809WczgkE44HWvtr/gnNDqMF34pkgSaAI0SCTaygSMVKgn1wp49M18ZaReQ2N35s0C3CgDEbdCcg8+1fanwd/a5+HPw58FaXpUuhavpeolmur65sIo5BNP"
    "u4cksC2EOFB4XBHGc142dxrVcJKlQhzNm1K3Om3Y+t5NUtfEkEOsIYLWyvFaeTLfu4gCfMBPTA2sc1z/AIH8RT+LND1LVXSaOC61ed4Fmz/qtqCPbnqCuOnG"
    "c1856/8AtreFdQ8OeLdKsLO9tEu5TDpRSDy1jgkG2d3UE7Tgs2wZyXb1Fb/w+/bS+GWn+BPDujanD4gsLvTrJLaSK2sUliZw2WYNvyck5GR2xxX5XiOGsXTw"
    "taNGleU5K3kt/wAz14YyLnBzekUdb+1LeQ6Z8OdJvbjz/KttbEm2CQpu/wBHfg46jpxXrNrq1r4msLXWbB1k0/UYVvLeSM5Uo4z19jlT/u18rfHf9rL4Y+O/"
    "hrrnhrRrXWNW1G48v7JJqFmLVbWXp9oVgzElRkbeN26uS+B/7W1v8KvCVl4a1jQW1nTFkknE1jJ5NzbOcDIByrKwGSDwPrXTU4cxeKyanR5LVYN6PqmJYyEM"
    "RKa2Z6f4ejaP9tfUNylU/suRlLDA5VhkV7/d31tpul3up3LKLDTreS5nMj7V2qpJBPbJwBn1FeCx/tO/CKXxtc+JX0rWodR/s4WSXrWMZlZQ2cBPMwDtO3cc"
    "ZA6DqeD+Of7Vui/EXSr3w1o1mYvC0kQSR5mC3F04yRIwGMAHChCD1YknitsTkmLzLF4bmhywhFKTfddiYYmFGE0t2zuf2RkuE8AfEjxBcRutxd3NxL9okJB/"
    "49pXUKf7oXaR9TWp+zN8cbf4jWGp+GNRknGuacxuLJp2cmW1+VTGuf7pywHoWrhPgx+1/wCA/hd8KdA0C+h1wanpQIl+x28c6XMx+9PudhwRgbG6BQOleUS/"
    "H/RtD+MUnjXwRpZ0S1853ispVyo5zgqOAGJJKjgdB2r6CvlU8weJpVqdlL4X6Kxx06vsnCcXsfd3i0zN4I8U2kMZma90m6hMQfaG/dkj9R/P1r4i/Zv8KP4t"
    "+NWmagIBPbWEzXMkbQmWMFUYqjcYXcwCgnGCRXu3if8Aa/8AAeseDr+K0sNdi1LULOSOGB7QLbo0iMv+tDE7QT/dzyPTNct+zv8AtH+A/h14Jm0bWtI1Pw5q"
    "U05lurjS7YXS3mD+6y7EMAg4Cn0z1NeVk+Bx+V5bXoum3Juy+fU2xFWnXrKV9D3bSvEHg3UvidqttaatLF4wtg1lcWd5uSN5AGXKtt2sSo3hQcg9RXbRssci"
    "lgWXJVljyCVPDj64J5r4yn+OPw9m+Pkfi+fw9cjR1vftryGLEtxOYNhnaFSVV1cK4AOCck9TXrV3+2r8K4LiPP8AwlV07nfJMumRjaSSTnMgzz6V4ma8OYp1"
    "KNXB03dJN+THgZUMKqkf5meU/CrwzP4U/bI13Sfscq2VlczwQzeSfLOCQrK2MEkY5HWvsOJGR4yoJYEEY657Y96+fLn9s74Yy+OoNQ/sjUJLWPTntzqJtI1u"
    "zNuGxvK37QAhZd2d3tzmtWx/bZ+GCahDhPE8bKd4lSxicIRyOkvqK6s7y3Mszq0KsaL91K+vU6MJiKVGM1J7no/wu1i71v4Z6K+pIyXtu11YSO2dzrHMdjEn"
    "knDY/wCA1b8XXNt4Sk0HxpLYG7m0+5a3uJ4rcOyW5QB5JZMEgLvByecKR2rwzSP23vCa67rNtqmiX+l+Hprt7vT54RG06u4Bl8yPcFwz5YBT8pJ6infEn9sj"
    "wXf+B9R03wz9s1W+1KF7eT7bbGCCJWXDEgElzj5e2M55qIZLmFPNXVjS/dz310s/8gdelLDqDeqPo1VXevlTrcRZGyaNw6yL1DAjggjBB9DX5+fGzR729+L/"
    "AIvjht5w8d6S7pbu7RISctgDqQrY9cH0NeofBv8AbHTS/DcOgeObOS+S1VYLLVNNULcCLACo6H5X29AcrgADtXpEP7RXwq8LeKtV8RixuGvL0RSDUYbNTO8w"
    "DKx2lzsLZB3YPU4A7+pk+XYzIsRX/d86l8NupjicRHExgtrHzf8AEb4A3/wn0rwHqVxdtLPq0pM1gtsUNowCSCMvn94wWRQeAQcivvwktDZkqUP2S3ypGCv7"
    "peK+DviP+0DP8QvijZ63exS3fhqxvYlg0qBpd0dssoMhz0Mjrks3GelfS4/bI+FWrSyT3E/iW3uZv3rxPpKfKxJ+X7/bjnpT4my/MMxwtGEYc0k23bp2Q8FV"
    "p0pylJ2RtX73KftE2kJuZktP+EXjfyPMIjLFWGdvQkgAe9djrP8AyA9Q/wCvaX/0A15Teftr/COLU40aHUppPINp/arWC/aYo+CIdhOCu7OGDZBHTBzViD9s"
    "L4V72eA+I5Zo1aQRNp0QD7ecfNJjn34r57GZNmGIqYeSotciSfnY6aWJpU41E5b3Pzz8eKVngBBHzOefword+OGv6L4i8V3N5oVp/Z1hcXc9zHZNjfCJCGwd"
    "vy49ACQOQOlFfulG/s1dWPnnuedwoHkAZto9cZro45baKNRI4YkA5xXM05nZ8ZOcDFaNXC51UGl29+5cXixBuQuzOP1pqaXbo4P2zOP+mf8A9eubivJYQAjA"
    "Y9qj3t/eP51PK+47nQtpEMU5lW7DbuxTGP1r1z4E6x4P0zXifF80f2AI6xCWAyQCUqNrXG3D+UCBkA9SO2a8DWZ16HP15q5Za1cWBJjWB85yJYg4I9DntWVW"
    "k6sHBsqEuSSl2Pu9fH/wjsUW1+2/DlfKATa2h3jEcdMljn65qSDxH8KETfDL8M0Ud10W7B/9Cr4M/tq4/uxY9Nn/ANehdZlVcCG2x/1zP+NeC8ji/wDl7L7z"
    "3FnErW9jD/wE+9P+Eo+FIYsbj4ZhvVtHu+f1pH8S/CZzlpvhm3uNHu/8a+CJ9Re4UB4YODn5UI/rRDqTwJtWG3Izn5oyf60f2HH/AJ+z+8P7Yl/z5h/4CfdK"
    "v8F7yeed4Phy0ine7/2TqABPqMSfyqZvE/wWlOWPw73H+JtI1AY/8fr4VXXLlJFkjEUbgYBRSOPTrTX1eaRCpSEA+keP61osmX/P2X3ny+KoLFVnWcnG/ROy"
    "+4+8rXxJ8G1ZvKk+G2cc50i/P83NSP4m+Du75m+GhP8A2CL3/wCKr4EXUJF6Kn5H/Gka+kc5IQfQH/Gn/Y6/5+y+85PqC/5+S+8++/7f+DErKXb4ZHBBydGu"
    "+P1rf0rV/wBnvSI2mWx8EKJCZHlOj3jK5PdSzHj2FfnYmszoqqI7cgDHzRZz9eau3PjG/ubOO28qyhjjGFMFqqMPxrmq5Cqqs6016SPRwcfqcnJPmv8Azan6"
    "Jx/ET9nksyrb+CGYdQNDujj9aZL8Qv2eQ/zQ+B0OOjaJdivzdOqXRYt5pBPUqAP5VPD4gvoE2rMCM5yy5Nc3+rNP/oIqf+BHr/X3/JH7j9GE8b/s7SSLiPwE"
    "GYjDNpF0ACT354rP8T+K/wBnt5EBTwLc3cONj22k3sqYz32Ng1+ebavO7EssTEnJynX9adBrU9oS1vFb28naSKLDD6HNbUuHo0aiqKvN26OWhjVxbrU3TcEr"
    "9Utfkff76x8EdXVHmi+HpCj5fN0O9jP6tUYvvgfC+yKP4cg9f+QRef8AxdfA1z4gvrxEWaUSbCSCc55/Goo9WuIzx5Z/3lz/ADr3lhbdT5t5dF/8vJfefe8n"
    "if4H7ykkXw9Yg7c/2Jekf+hVcsb/AOC17KyQJ8OWYxscf2PeDIA5HLf/AF6/P46zdNn5k59EApYtcvICCkgBByCcnB7Hk9abwt1uH9nR/wCfsvvPS/2jx4RH"
    "jMN4LEA0WUebCLbcEAKIGCq3zKu8PhWJIFFeZajq0+qEGcRAhi+Y4whJOM5x9P50V1048kVFs9SK5VYpUUUVoUFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFF"
    "FFABRRRQAUUUUAFFFFABRRRQAUUUUAf/2Q=="
)
import base64 as _b64
LOGO_BYTES = _b64.b64decode(LOGO_B64)
N8N_CHAT_URL = os.getenv("N8N_CHAT_URL", "https://quiver-ointment-evacuee.ngrok-free.dev/webhook/chat-bot").strip()
REF_TABLE_MAP = {
    "lob":    ("ref_lob", "name"),
    "biz":    ("ref_business_line", "name"),
    "pm":     ("ref_project_managers", "name"),
    "status": ("ref_project_status", "name"),
    "doccat": ("ref_doc_category", "name"),
    "expcat": ("ref_expense_category", "name"),
}

# ═══════════════════════════════════════════════════════════════
#  AUTH
# ═══════════════════════════════════════════════════════════════

def _hash_pw(pw: str) -> str:
    return hashlib.sha256(pw.encode()).hexdigest()

def fetch_all_users() -> list[dict]:
    try:
        with ENGINE.connect() as conn:
            rows = conn.execute(text(
                "SELECT id, username, display_name, role, pm_name, created_at "
                "FROM users ORDER BY id ASC"
            )).fetchall()
        return [dict(zip(["id","username","display_name","role","pm_name","created_at"], r)) for r in rows]
    except Exception as e:
        print(f"[fetch_all_users] {e}")
        return []


def create_user(username: str, password: str, display_name: str = "", role: str = "user", pm_name: str = "") -> tuple[bool, str]:
    username = (username or "").strip()
    password = password or ""
    if not username or not password:
        return False, "Username and password are required."
    if len(password) < 6:
        return False, "Password must be at least 6 characters."
    try:
        with ENGINE.begin() as conn:
            exists = conn.execute(
                text("SELECT id FROM users WHERE username=:u"), {"u": username}
            ).fetchone()
            if exists:
                return False, f'Username "{username}" already exists.'
            conn.execute(text("""
                INSERT INTO users (username, password_hash, display_name, role, pm_name, created_at)
                VALUES (:u, :ph, :dn, :r, :pm, :ca)
            """), {
                "u": username, "ph": _hash_pw(password),
                "dn": (display_name or username).strip(),
                "r": role or "user",
                "pm": (pm_name or "").strip(),
                "ca": datetime.now().isoformat(),
            })
        return True, f'User "{username}" created successfully.'
    except Exception as e:
        print(f"[create_user] {e}")
        return False, f"Could not create user: {e}"


def delete_user_account(user_id: int) -> bool:
    try:
        with ENGINE.begin() as conn:
            conn.execute(text("DELETE FROM users WHERE id=:id"), {"id": user_id})
        return True
    except Exception as e:
        print(f"[delete_user_account] {e}")
        return False


def update_user_role(user_id: int, role: str) -> bool:
    try:
        with ENGINE.begin() as conn:
            conn.execute(text("UPDATE users SET role=:r WHERE id=:id"), {"r": role, "id": user_id})
        return True
    except Exception as e:
        print(f"[update_user_role] {e}")
        return False

def verify_user(username: str, password: str):
    try:
        with ENGINE.connect() as conn:
            row = conn.execute(
                text("SELECT id,username,password_hash,display_name,role,pm_name "
                     "FROM users WHERE username=:u"), {"u": username}
            ).fetchone()
        if row and row[2] == _hash_pw(password):
            return {"id": row[0], "username": row[1],
                    "display_name": row[3] or row[1], "role": row[4],
                    "pm_name": row[5] or ""}
    except Exception:
        pass
    return None

def create_session(user: dict) -> str:
    token = secrets.token_hex(24)
    SESSIONS[token] = {**user, "expires": datetime.now() + timedelta(seconds=SESSION_LIFETIME)}
    return token

def destroy_session(token: str):
    SESSIONS.pop(token, None)
def reset_user_password(username: str, new_password: str) -> bool:
    try:
        with ENGINE.begin() as conn:
            row = conn.execute(text("SELECT id FROM users WHERE username=:u"), {"u": username}).fetchone()
            if not row:
                return False
            conn.execute(text("UPDATE users SET password_hash=:ph WHERE username=:u"),
                         {"ph": _hash_pw(new_password), "u": username})
        return True
    except Exception:
        return False

def generate_pm_logins() -> tuple[int, list[str]]:
    """Creates a login for every Project Manager who doesn't already have one.
    Username = slugified PM name; default password = 'Monarch@123'."""
    created = []
    pms = get_project_managers()
    existing = set()
    try:
        with ENGINE.connect() as conn:
            rows = conn.execute(text(
                "SELECT pm_name FROM users WHERE pm_name IS NOT NULL AND pm_name <> ''"
            )).fetchall()
        existing = {r[0].strip().lower() for r in rows if r[0]}
    except Exception:
        pass

    count = 0
    for pm in pms:
        if pm.strip().lower() in existing:
            continue
        base = re.sub(r'[^a-z0-9]+', '.', pm.strip().lower()).strip('.') or f"pm{count}"
        username, n = base, 1
        while True:
            with ENGINE.connect() as conn:
                exists = conn.execute(text("SELECT id FROM users WHERE username=:u"),
                                       {"u": username}).fetchone()
            if not exists:
                break
            n += 1
            username = f"{base}{n}"
        default_pw = "Monarch@123"
        ok, _ = create_user(username, default_pw, display_name=pm, role="pm", pm_name=pm)
        if ok:
            created.append(f"{pm} → {username} / {default_pw}")
            count += 1
    return count, created
def _get_session_token(handler) -> str | None:
    for part in handler.headers.get("Cookie", "").split(";"):
        p = part.strip()
        if p.startswith(SESSION_COOKIE + "="):
            return p.split("=", 1)[1].strip()
    return None

def get_session_user(handler) -> str | None:
    token = _get_session_token(handler)
    if not token: return None
    sess  = SESSIONS.get(token)
    if not sess: return None
    if sess["expires"] < datetime.now():
        SESSIONS.pop(token, None)
        return None
    return sess.get("display_name") or sess.get("username")

def get_session_role(handler) -> str | None:
    """Returns the role ('user', 'admin', 'superadmin', etc.) of the currently
    logged-in session, or None if not logged in / session expired."""
    token = _get_session_token(handler)
    if not token: return None
    sess  = SESSIONS.get(token)
    if not sess: return None
    if sess["expires"] < datetime.now():
        SESSIONS.pop(token, None)
        return None
    return sess.get("role")


def is_admin_role(role: str | None) -> bool:
    return (role or "").lower() in ("admin", "superadmin")


def get_session_pm_name(handler) -> str | None:
    token = _get_session_token(handler)
    if not token: return None
    sess = SESSIONS.get(token)
    if not sess: return None
    if sess["expires"] < datetime.now():
        SESSIONS.pop(token, None)
        return None
    return sess.get("pm_name") or ""


def get_session_active_filter(handler) -> bool:
    """Whether the global 'Active Projects only' toggle is currently on for
    this user's session. Stored on the session itself (not a cookie/JS
    state) so it stays in effect across every page until switched off."""
    token = _get_session_token(handler)
    if not token: return False
    sess = SESSIONS.get(token)
    if not sess: return False
    return bool(sess.get("active_only", False))


def toggle_session_active_filter(handler) -> bool:
    """Flips the global Active-only filter for this session and returns the new value."""
    token = _get_session_token(handler)
    if not token or token not in SESSIONS:
        return False
    new_val = not bool(SESSIONS[token].get("active_only", False))
    SESSIONS[token]["active_only"] = new_val
    return new_val


def is_pm_role(role: str | None) -> bool:
    return (role or "").lower() == "pm"

def is_head_role(role: str | None) -> bool:
    return (role or "").lower() == "head"

def is_procurement_role(role: str | None) -> bool:
    return (role or "").lower() == "procurement"

def is_accounts_role(role: str | None) -> bool:
    return (role or "").lower() == "accounts"

def is_user_role(role: str | None) -> bool:
    """The generic 'user' role (assigned via User Management) — acts as the
    approver for the Procurement and Accounts stages on the Payment Approvals
    screen. Final sign-off still belongs to 'director'."""
    return (role or "").lower() == "user"

def is_booking_role(role: str | None) -> bool:
    """PM or Head — the two roles allowed to create Weekly Payment Bookings (Screen 1)."""
    return (role or "").lower() in ("pm", "head")


# ═══════════════════════════════════════════════════════════════
#  DATABASE — SCHEMA INIT
# ═══════════════════════════════════════════════════════════════
def migrate_db():
    """Add any missing columns to existing tables."""
    migrations = [
        ("users", "pm_name", "ALTER TABLE users ADD COLUMN IF NOT EXISTS pm_name TEXT"),
        ("projects", "variation_amt",                "ALTER TABLE projects ADD COLUMN IF NOT EXISTS variation_amt REAL DEFAULT 0"),
        ("projects", "revised_order_value",          "ALTER TABLE projects ADD COLUMN IF NOT EXISTS revised_order_value REAL DEFAULT 0"),
        ("projects", "hier_business_line",           "ALTER TABLE projects ADD COLUMN IF NOT EXISTS hier_business_line TEXT"),
        ("projects", "parent_project",               "ALTER TABLE projects ADD COLUMN IF NOT EXISTS parent_project TEXT"),
        ("projects", "child_project",                "ALTER TABLE projects ADD COLUMN IF NOT EXISTS child_project TEXT"),
        ("projects", "baseline_eac",                 "ALTER TABLE projects ADD COLUMN IF NOT EXISTS baseline_eac REAL DEFAULT 0"),
        ("projects", "baseline_margin",              "ALTER TABLE projects ADD COLUMN IF NOT EXISTS baseline_margin REAL DEFAULT 0"),
        ("projects", "rev_latest_eac",               "ALTER TABLE projects ADD COLUMN IF NOT EXISTS rev_latest_eac REAL DEFAULT 0"),
        ("projects", "rev_latest_margin",            "ALTER TABLE projects ADD COLUMN IF NOT EXISTS rev_latest_margin REAL DEFAULT 0"),
        ("projects", "physical_completion",          "ALTER TABLE projects ADD COLUMN IF NOT EXISTS physical_completion REAL DEFAULT 0"),
        ("projects", "billed_amount",                "ALTER TABLE projects ADD COLUMN IF NOT EXISTS billed_amount REAL DEFAULT 0"),
        ("projects", "last_billed_date",             "ALTER TABLE projects ADD COLUMN IF NOT EXISTS last_billed_date TEXT"),
        ("projects", "collection_till_date",         "ALTER TABLE projects ADD COLUMN IF NOT EXISTS collection_till_date REAL DEFAULT 0"),
        ("projects", "itd_cost",                     "ALTER TABLE projects ADD COLUMN IF NOT EXISTS itd_cost REAL DEFAULT 0"),
        ("projects", "bill_submission_pending_date", "ALTER TABLE projects ADD COLUMN IF NOT EXISTS bill_submission_pending_date TEXT"),
        ("projects", "project_type",                 "ALTER TABLE projects ADD COLUMN IF NOT EXISTS project_type TEXT DEFAULT 'parent'"),
        ("projects", "parent_id",                    "ALTER TABLE projects ADD COLUMN IF NOT EXISTS parent_id INTEGER"),
        ("projects", "parent_id_to_text",             "ALTER TABLE projects ALTER COLUMN parent_id TYPE TEXT"),
        ("projects", "folder_path", "ALTER TABLE projects ADD COLUMN IF NOT EXISTS folder_path TEXT"),
        ("projects", "end_date",                     "ALTER TABLE projects ADD COLUMN IF NOT EXISTS end_date TEXT"),
        ("projects", "extension_months",             "ALTER TABLE projects ADD COLUMN IF NOT EXISTS extension_months REAL DEFAULT 0"),
        ("projects", "extended_end_date",           "ALTER TABLE projects ADD COLUMN IF NOT EXISTS extended_end_date TEXT"),
        # ── CO.No is now the linking key for the Expenditures feature (instead of the
        #    internal numeric project_id). These columns are additive — project_id is
        #    left in place for now so nothing else breaks, but the app no longer reads it.
        ("expenditures",   "co_no", "ALTER TABLE expenditures ADD COLUMN IF NOT EXISTS co_no TEXT"),
        ("erp_updates",    "co_no", "ALTER TABLE erp_updates ADD COLUMN IF NOT EXISTS co_no TEXT"),
        ("expense_entries","co_no", "ALTER TABLE expense_entries ADD COLUMN IF NOT EXISTS co_no TEXT"),
        # ── payment_bookings existed on some deployments from an earlier, narrower
        #    schema (pre-dating the Weekly Payment Booking / Payment Approvals
        #    feature set). "CREATE TABLE IF NOT EXISTS" is a no-op on an existing
        #    table, so backfill every column the app now expects. ──
        ("payment_bookings", "booking_uid",         "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS booking_uid TEXT UNIQUE"),
        ("payment_bookings", "project_id",           "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS project_id INTEGER REFERENCES projects(id) ON DELETE SET NULL"),
        ("payment_bookings", "is_indirect",          "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS is_indirect BOOLEAN DEFAULT FALSE"),
        ("payment_bookings", "indirect_dept",        "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS indirect_dept TEXT"),
        ("payment_bookings", "location",              "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS location TEXT"),
        ("payment_bookings", "project_label",        "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS project_label TEXT"),
        ("payment_bookings", "pm_username",          "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS pm_username TEXT"),
        ("payment_bookings", "pm_name",               "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS pm_name TEXT"),
        ("payment_bookings", "payee_name",            "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS payee_name TEXT"),
        ("payment_bookings", "payment_category",      "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS payment_category TEXT"),
        ("payment_bookings", "requested_amount",      "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS requested_amount REAL DEFAULT 0"),
        ("payment_bookings", "invoice_date",          "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS invoice_date TEXT"),
        ("payment_bookings", "justification",         "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS justification TEXT"),
        ("payment_bookings", "period_start",          "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS period_start TEXT"),
        ("payment_bookings", "period_end",            "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS period_end TEXT"),
        ("payment_bookings", "overall_status",        "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS overall_status TEXT DEFAULT 'pending'"),
        ("payment_bookings", "procurement_status",    "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS procurement_status TEXT DEFAULT 'pending'"),
        ("payment_bookings", "procurement_note",      "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS procurement_note TEXT"),
        ("payment_bookings", "procurement_by",        "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS procurement_by TEXT"),
        ("payment_bookings", "procurement_at",        "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS procurement_at TEXT"),
        ("payment_bookings", "accounts_status",       "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS accounts_status TEXT DEFAULT 'pending'"),
        ("payment_bookings", "accounts_note",         "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS accounts_note TEXT"),
        ("payment_bookings", "accounts_by",           "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS accounts_by TEXT"),
        ("payment_bookings", "accounts_at",           "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS accounts_at TEXT"),
        ("payment_bookings", "director_status",       "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS director_status TEXT DEFAULT 'pending'"),
        ("payment_bookings", "director_note",         "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS director_note TEXT"),
        ("payment_bookings", "director_by",           "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS director_by TEXT"),
        ("payment_bookings", "director_at",           "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS director_at TEXT"),
        ("payment_bookings", "device_id",             "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS device_id TEXT"),
        ("payment_bookings", "digitally_signed_by",   "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS digitally_signed_by TEXT"),
        ("payment_bookings", "archived",              "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS archived BOOLEAN DEFAULT FALSE"),
        ("payment_bookings", "created_at",            "ALTER TABLE payment_bookings ADD COLUMN IF NOT EXISTS created_at TEXT"),
        # ── Older deployments of this table also carry a legacy "payment_id"
        #    NOT NULL column from before this feature existed. The app never
        #    populates it, so relax the constraint rather than fail every insert.
        #    (Skipped harmlessly if the column doesn't exist.) ──
        ("payment_bookings", "payment_id_nullable",   "ALTER TABLE payment_bookings ALTER COLUMN payment_id DROP NOT NULL"),
        # ── payment_booking_docs has the same legacy-table problem. ──
        ("payment_booking_docs", "booking_id",  "ALTER TABLE payment_booking_docs ADD COLUMN IF NOT EXISTS booking_id INTEGER REFERENCES payment_bookings(id) ON DELETE CASCADE"),
        ("payment_booking_docs", "doc_type",    "ALTER TABLE payment_booking_docs ADD COLUMN IF NOT EXISTS doc_type TEXT"),
        ("payment_booking_docs", "file_name",   "ALTER TABLE payment_booking_docs ADD COLUMN IF NOT EXISTS file_name TEXT"),
        ("payment_booking_docs", "file_path",   "ALTER TABLE payment_booking_docs ADD COLUMN IF NOT EXISTS file_path TEXT"),
        ("payment_booking_docs", "uploaded_at", "ALTER TABLE payment_booking_docs ADD COLUMN IF NOT EXISTS uploaded_at TEXT"),
    ]
    # Each migration runs in its OWN transaction. PostgreSQL aborts the whole
    # transaction after one failed statement, which previously made every later
    # migration (including parent_id -> TEXT) silently fail, so saving a new
    # project broke on existing databases.
    for table, col, sql in migrations:
        try:
            with ENGINE.begin() as mconn:
                mconn.execute(text(sql))
            print(f"[migrate] Ensured column: {table}.{col}")
        except Exception as e:
            print(f"[migrate] Skipped {table}.{col}: {str(e).splitlines()[0]}")

    with ENGINE.begin() as conn:

        # ── Generic safety net: some deployments have a payment_bookings table
        #    inherited from an even earlier schema version, with stray NOT NULL
        #    columns (like "payment_id" above) that this app never populates and
        #    doesn't know the name of in advance. Rather than hunt these down
        #    one error at a time, find any such column automatically and relax
        #    the constraint so inserts stop failing on unrelated legacy fields. ──
        KNOWN_PAYMENT_BOOKING_COLS = {
            "id", "booking_uid", "project_id", "is_indirect", "indirect_dept",
            "project_label", "pm_username", "pm_name", "payee_name",
            "payment_category", "requested_amount", "invoice_date",
            "justification", "period_start", "period_end", "overall_status",
            "procurement_status", "procurement_note", "procurement_by",
            "procurement_at", "accounts_status", "accounts_note", "accounts_by",
            "accounts_at", "director_status", "director_note", "director_by",
            "director_at", "device_id", "digitally_signed_by", "archived",
            "created_at",
        }
        KNOWN_PAYMENT_BOOKING_DOC_COLS = {
            "id", "booking_id", "doc_type", "file_name", "file_path", "uploaded_at",
        }
        for tbl, known_cols in (
            ("payment_bookings", KNOWN_PAYMENT_BOOKING_COLS),
            ("payment_booking_docs", KNOWN_PAYMENT_BOOKING_DOC_COLS),
        ):
            try:
                stray_cols = conn.execute(text("""
                    SELECT column_name FROM information_schema.columns
                    WHERE table_schema = current_schema()
                      AND table_name = :tbl
                      AND is_nullable = 'NO'
                      AND column_default IS NULL
                """), {"tbl": tbl}).fetchall()
                for (col_name,) in stray_cols:
                    if col_name in known_cols:
                        continue
                    try:
                        conn.execute(text(
                            f'ALTER TABLE {tbl} ALTER COLUMN "{col_name}" DROP NOT NULL'
                        ))
                        print(f"[migrate] Relaxed stray NOT NULL: {tbl}.{col_name}")
                    except Exception as e:
                        print(f"[migrate] Could not relax {tbl}.{col_name}: {e}")
            except Exception as e:
                print(f"[migrate] Stray-column sweep skipped for {tbl}: {e}")

        # ── Backfill co_no on existing rows from the linked project, one-time per row.
        #    Safe to re-run: only touches rows where co_no is still NULL. ──
        backfills = [
            ("expenditures",    "UPDATE expenditures SET co_no = p.co_no "
                                 "FROM projects p WHERE expenditures.project_id = p.id "
                                 "AND expenditures.co_no IS NULL"),
            ("erp_updates",     "UPDATE erp_updates SET co_no = p.co_no "
                                 "FROM projects p WHERE erp_updates.project_id = p.id "
                                 "AND erp_updates.co_no IS NULL"),
            ("expense_entries", "UPDATE expense_entries SET co_no = p.co_no "
                                 "FROM projects p WHERE expense_entries.project_id = p.id "
                                 "AND expense_entries.co_no IS NULL"),
        ]
        for table, sql in backfills:
            try:
                result = conn.execute(text(sql))
                print(f"[migrate] Backfilled co_no on {table}: {result.rowcount} row(s)")
            except Exception as e:
                print(f"[migrate] Backfill skipped for {table}: {e}")

        # ── Clean up stale/duplicate LOB casing (e.g. 'Geotech' -> 'GEOTECH')
        try:
            conn.execute(text("UPDATE projects SET business_line = 'GEOTECH' WHERE LOWER(TRIM(business_line)) = 'geotech'"))
            conn.execute(text("UPDATE projects SET hier_business_line = 'GEOTECH' WHERE LOWER(TRIM(hier_business_line)) = 'geotech'"))
            conn.execute(text("UPDATE expenditures SET lob = 'GEOTECH' WHERE LOWER(TRIM(lob)) = 'geotech'"))
            conn.execute(text("DELETE FROM ref_lob WHERE LOWER(TRIM(name)) = 'geotech' AND name <> 'GEOTECH'"))
            conn.execute(text("DELETE FROM ref_business_line WHERE LOWER(TRIM(name)) = 'geotech' AND name <> 'GEOTECH'"))
        except Exception as e:
            print(f"[migrate] LOB cleanup: {e}")

def init_db():
    with ENGINE.begin() as conn:

        conn.execute(text('CREATE SCHEMA IF NOT EXISTS "monarch_db"'))
        
        migrate_db()

        result = conn.execute(text("""
            SELECT schema_name
            FROM information_schema.schemata
            WHERE schema_name='monarch_db'
        """)).fetchall()

        print("Schema Check:", result)

        conn.execute(text('SET search_path TO "monarch_db"'))

        print(
            "Current Schema:",
            conn.execute(text("SELECT current_schema()")).scalar()
        )

        # CREATE TABLE statements here
    """Create all tables and seed reference data if empty."""
    with ENGINE.begin() as conn:

        conn.execute(text(f"""
            CREATE TABLE IF NOT EXISTS {SCHEMA_NAME}.users (
                id SERIAL PRIMARY KEY,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                display_name TEXT,
                role TEXT DEFAULT 'user',
                created_at TEXT
            )
        """))
        conn.execute(text("""
    CREATE TABLE IF NOT EXISTS erp_updates (
        id          SERIAL PRIMARY KEY,
        project_id  INTEGER REFERENCES projects(id) ON DELETE CASCADE,
        co_no       TEXT,
        category    TEXT NOT NULL,
        amount      REAL DEFAULT 0,
        created_by  TEXT,
        created_at  TEXT
    )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS ref_lob (
                id     SERIAL PRIMARY KEY,
                name   TEXT UNIQUE NOT NULL,
                active BOOLEAN DEFAULT TRUE
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS lob_ebf_settings (
                lob        TEXT PRIMARY KEY,
                ebf_mult   REAL DEFAULT 1.35,
                updated_at TEXT
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS ref_project_managers (
                id           SERIAL PRIMARY KEY,
                name TEXT UNIQUE NOT NULL,
                email        TEXT,
                active       BOOLEAN DEFAULT TRUE
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS ref_project_status (
                id     SERIAL PRIMARY KEY,
                name   TEXT UNIQUE NOT NULL,
                color  TEXT DEFAULT '#43474e',
                active BOOLEAN DEFAULT TRUE
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS ref_business_line (
                id     SERIAL PRIMARY KEY,
                name   TEXT UNIQUE NOT NULL,
                active BOOLEAN DEFAULT TRUE
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS ref_doc_category (
                id     SERIAL PRIMARY KEY,
                name   TEXT UNIQUE NOT NULL,
                active BOOLEAN DEFAULT TRUE
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS ref_expense_category (
                id       SERIAL PRIMARY KEY,
                name     TEXT UNIQUE NOT NULL,
                color_bg TEXT DEFAULT '#f5f5f5',
                color_fg TEXT DEFAULT '#374151',
                active   BOOLEAN DEFAULT TRUE
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS ref_clients (
                id     SERIAL PRIMARY KEY,
                name   TEXT UNIQUE NOT NULL,
                active BOOLEAN DEFAULT TRUE
            )"""))

        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS projects (
                id                   SERIAL PRIMARY KEY,
                co_no                TEXT,
                project_name         TEXT,
                business_line        TEXT,
                project_manager      TEXT,
                client_name          TEXT,
                order_number         TEXT,
                order_issue_date     TEXT,
                end_date             TEXT,
                extension_months     REAL DEFAULT 0,
                extended_end_date    TEXT,
                order_value          REAL DEFAULT 0,
                variation_amt        REAL DEFAULT 0,
                revised_order_value  REAL DEFAULT 0,
                hier_business_line   TEXT,
                parent_project       TEXT,
                child_project        TEXT,
                baseline_eac         REAL DEFAULT 0,
                baseline_margin      REAL DEFAULT 0,
                rev_latest_eac       REAL DEFAULT 0,
                rev_latest_margin    REAL DEFAULT 0,
                project_status       TEXT DEFAULT 'Active',
                physical_completion  REAL DEFAULT 0,
                billed_amount        REAL DEFAULT 0,
                last_billed_date     TEXT,
                collection_till_date REAL DEFAULT 0,
                itd_cost             REAL DEFAULT 0,
                bill_submission_pending_date TEXT,
                project_type         TEXT DEFAULT 'parent',
                parent_id            TEXT,
                created_at           TEXT
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS documents (
                id           SERIAL PRIMARY KEY,
                project_id   INTEGER REFERENCES projects(id) ON DELETE CASCADE,
                doc_name     TEXT,
                doc_category TEXT,
                doc_link     TEXT,
                created_at   TEXT
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS expenditures (
                id              SERIAL PRIMARY KEY,
                project_id      INTEGER REFERENCES projects(id) ON DELETE CASCADE,
                co_no           TEXT,
                lob             TEXT,
                data_as_of      TEXT,
                erp_update_date TEXT,
                itd_cogs        REAL DEFAULT 0,
                itd_djcs        REAL DEFAULT 0,
                itd_equip       REAL DEFAULT 0,
                itd_manpower    REAL DEFAULT 0,
                itd_vehicles    REAL DEFAULT 0,
                itd_others      REAL DEFAULT 0,
                erp_cogs        REAL DEFAULT 0,
                erp_djcs        REAL DEFAULT 0,
                erp_equip       REAL DEFAULT 0,
                erp_manpower    REAL DEFAULT 0,
                erp_vehicles    REAL DEFAULT 0,
                erp_others      REAL DEFAULT 0,
                created_at      TEXT
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS expense_entries (
                id             SERIAL PRIMARY KEY,
                expenditure_id INTEGER,
                project_id     INTEGER REFERENCES projects(id) ON DELETE CASCADE,
                co_no          TEXT,
                category       TEXT,
                description    TEXT,
                quantity       REAL DEFAULT 0,
                unit_cost      REAL DEFAULT 0,
                line_total     REAL DEFAULT 0,
                created_at     TEXT
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS variations (
                id             SERIAL PRIMARY KEY,
                project_id     INTEGER REFERENCES projects(id) ON DELETE CASCADE,
                child_co_no    TEXT,
                child_name     TEXT,
                parent_name    TEXT,
                original_value REAL DEFAULT 0,
                variation_amt  REAL DEFAULT 0,
                revised_value  REAL DEFAULT 0,
                variation_date TEXT,
                status         TEXT DEFAULT 'Pending',
                created_at     TEXT
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS billing_collection (
                id              SERIAL PRIMARY KEY,
                project_id      INTEGER REFERENCES projects(id) ON DELETE CASCADE,
                ra_no           TEXT,
                billing_date    TEXT,
                billed_amt      REAL DEFAULT 0,
                collection_date TEXT,
                collection_amt  REAL DEFAULT 0,
                created_at      TEXT
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS payment_bookings (
                id                   SERIAL PRIMARY KEY,
                booking_uid          TEXT UNIQUE,
                project_id           INTEGER REFERENCES projects(id) ON DELETE SET NULL,
                is_indirect          BOOLEAN DEFAULT FALSE,
                indirect_dept        TEXT,
                location             TEXT,
                project_label        TEXT,
                pm_username          TEXT,
                pm_name              TEXT,
                payee_name           TEXT,
                payment_category     TEXT,
                requested_amount     REAL DEFAULT 0,
                invoice_date         TEXT,
                justification        TEXT,
                period_start         TEXT,
                period_end           TEXT,
                overall_status       TEXT DEFAULT 'pending',
                procurement_status   TEXT DEFAULT 'pending',
                procurement_note     TEXT,
                procurement_by       TEXT,
                procurement_at       TEXT,
                accounts_status      TEXT DEFAULT 'pending',
                accounts_note        TEXT,
                accounts_by          TEXT,
                accounts_at          TEXT,
                director_status      TEXT DEFAULT 'pending',
                director_note        TEXT,
                director_by          TEXT,
                director_at          TEXT,
                device_id            TEXT,
                digitally_signed_by  TEXT,
                archived             BOOLEAN DEFAULT FALSE,
                created_at           TEXT
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS payment_booking_docs (
                id           SERIAL PRIMARY KEY,
                booking_id   INTEGER REFERENCES payment_bookings(id) ON DELETE CASCADE,
                doc_type     TEXT,
                file_name    TEXT,
                file_path    TEXT,
                uploaded_at  TEXT
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS user_devices (
                id          SERIAL PRIMARY KEY,
                username    TEXT UNIQUE,
                device_id   TEXT,
                first_seen  TEXT
            )"""))

        # ── Chat bot tables (n8n) ────────────────────────────────────────
        # These live in the PUBLIC schema (not monarch_db) because the n8n
        # workflow's Postgres nodes write there by default. IF NOT EXISTS
        # means this is a no-op on every restart after the first run — it
        # never touches or clears existing rows, so chat history persists.
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS public.chat_sessions (
                id         TEXT PRIMARY KEY,
                username   TEXT,
                title      TEXT,
                created_at TIMESTAMP DEFAULT now(),
                updated_at TIMESTAMP DEFAULT now()
            )"""))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS public.chat_messages (
                id         SERIAL PRIMARY KEY,
                session_id TEXT NOT NULL REFERENCES public.chat_sessions(id) ON DELETE CASCADE,
                who        TEXT NOT NULL,
                message    TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT now()
            )"""))
        conn.execute(text("""
            CREATE INDEX IF NOT EXISTS idx_chat_messages_session
            ON public.chat_messages(session_id)
        """))

        exists = conn.execute(text("SELECT id FROM users WHERE username='admin'")).fetchone()
        if not exists:
            conn.execute(text("""
                INSERT INTO users (username,password_hash,display_name,role,created_at)
                VALUES ('admin',:ph,'Admin','superadmin',:ca)
            """), {"ph": _hash_pw("monarch123"), "ca": datetime.now().isoformat()})

        _seed_ref(conn, "ref_lob", "name", [
            "Railways", "DPR-Highways", "Land Acquisition", "Locals",
            "Geotechnical", "Geospatial", "Project Management", "ELU preparation",
        ])
        _seed_ref(conn, "ref_project_managers", "name", [])
        _seed_ref(conn, "ref_project_status", "name", [
            "Active", "Closed", "Execution Closed", "Cancelled", "Hold",
        ])
        _seed_ref(conn, "ref_business_line", "name", [
            "Construction", "Engineering", "Surveying",
        ])
        _seed_ref(conn, "ref_doc_category", "name", [
            "Contract", "Technical Drawing", "Site Photo", "Permit", "Report", "Other",
        ])
        _seed_ref_expense(conn)
        _seed_ref(conn, "ref_clients", "name", [
            "National Highways Authority", "NHAI India", "Urban Dev Authority",
            "Apex Real Estate", "Water Resources Board",
        ])


def _seed_ref(conn, table: str, col: str, values: list):
    existing = conn.execute(text(f"SELECT {col} FROM {table}")).fetchall()
    existing_vals = {r[0] for r in existing}
    for v in values:
        if v not in existing_vals:
            conn.execute(text(f"INSERT INTO {table} ({col}) VALUES (:v)"), {"v": v})


def _seed_ref_expense(conn):
    existing = conn.execute(text("SELECT name FROM ref_expense_category")).fetchall()
    if existing: return
    defaults = [
        ("COGS",      "#e8edf7", "#00234b"),
        ("DJCS",      "#fdf3e0", "#775a19"),
        ("Equipment", "#eaf4eb", "#166534"),
        ("Manpower",  "#fde8e8", "#9b1c1c"),
        ("Vehicles",  "#f0e8fd", "#5b21b6"),
        ("Others",    "#f5f5f5", "#374151"),
    ]
    for name, bg, fg in defaults:
        conn.execute(text("""
            INSERT INTO ref_expense_category (name,color_bg,color_fg)
            VALUES (:n,:bg,:fg)
        """), {"n": name, "bg": bg, "fg": fg})


# ═══════════════════════════════════════════════════════════════
#  REFERENCE DATA FETCHERS
# ═══════════════════════════════════════════════════════════════

def _fetch_list(table: str, col: str = "name", order: str = "name") -> list[str]:
    try:
        with ENGINE.connect() as conn:
            rows = conn.execute(text(
                f"SELECT {col} FROM {table} WHERE active IS NOT FALSE ORDER BY {order}"
            )).fetchall()
        return [r[0] for r in rows]
    except Exception as e:
        print(f"[_fetch_list] {table}.{col} error: {e}")
        return []

def normalize_lob_name(name: str) -> str:
    s = str(name or "").strip()
    if s.lower() == "geotech":
        return "GEOTECH"
    return s

def get_lob_list() -> list[str]:
    raw = _fetch_list("ref_lob")
    seen = set()
    result = []
    for x in raw:
        norm = normalize_lob_name(x)
        if norm and norm.lower() not in seen:
            seen.add(norm.lower())
            result.append(norm)
    return result

def get_project_managers()   -> list[str]: return _fetch_list("ref_project_managers", "name")
def get_project_statuses()   -> list[str]: return _fetch_list("ref_project_status")
def get_business_lines()     -> list[str]: return _fetch_list("ref_business_line")
def get_doc_categories()     -> list[str]: return _fetch_list("ref_doc_category")
def get_expense_categories() -> list[str]: return _fetch_list("ref_expense_category")
def get_client_list()        -> list[str]: return _fetch_list("ref_clients")

def get_visible_lob_list() -> list[str]:
    """LOB list scoped to the logged-in PM's own projects when role == 'pm';
    otherwise the full master ref_lob list. Every LOB filter dropdown across
    the app uses this instead of get_lob_list() directly, so a PM only ever
    sees the Line(s) of Business tied to their own project(s), never the
    full company list."""
    role = (getattr(_req_ctx, "role", "") or "").lower()
    if role != "pm":
        return get_lob_list()
    pm_name = (getattr(_req_ctx, "pm_name", "") or "").strip().lower()
    if not pm_name:
        return []
    try:
        with ENGINE.connect() as conn:
            rows = conn.execute(text("""
                SELECT DISTINCT COALESCE(NULLIF(business_line,''), hier_business_line) AS lob
                FROM projects
                WHERE LOWER(TRIM(project_manager)) = :pm
            """), {"pm": pm_name}).fetchall()
        seen = set()
        lobs = []
        for r in rows:
            if r[0] and str(r[0]).strip():
                norm = normalize_lob_name(r[0])
                if norm.lower() not in seen:
                    seen.add(norm.lower())
                    lobs.append(norm)
        lobs.sort(key=lambda s: s.lower())
        return lobs
    except Exception as e:
        print(f"[get_visible_lob_list] {e}")
        return []

def get_expense_category_colors() -> dict:
    try:
        with ENGINE.connect() as conn:
            rows = conn.execute(text(
                "SELECT name,color_bg,color_fg FROM ref_expense_category WHERE active=TRUE"
            )).fetchall()
        return {r[0]: (r[1], r[2]) for r in rows}
    except Exception:
        return {}

def get_parent_project_options() -> list[dict]:
    """Returns parent projects as [{'id':.., 'co_no':.., 'label':'CO_NO-Project Name', 'name':'Project Name'}, ...].
    'co_no' is the authoritative key for linking a sub-project to its parent (co_no is unique per project)."""
    try:
        with ENGINE.connect() as conn:
            rows = conn.execute(text(
                "SELECT id, COALESCE(co_no,''), project_name FROM projects "
                "WHERE project_type='parent' ORDER BY project_name"
            )).fetchall()
        result = []
        for pid, co, name in rows:
            if not name:
                continue
            result.append({"id": pid, "co_no": co, "label": (f"{co}-{name}" if co else name), "name": name})
        return result
    except Exception:
        return []

def get_parent_project_names() -> list[str]:
    """Returns parent projects as 'CO_NO-Project Name' combined strings. Kept for backward compatibility —
    prefer get_parent_project_options() when you also need the co_no/id."""
    return [o["label"] for o in get_parent_project_options()]

def resolve_parent_by_co_no(co_no) -> dict:
    """Look up a parent project row strictly by co_no (co_no is unique per project). Returns {} if
    not found, not a parent, or co_no is blank."""
    co_no = (str(co_no).strip() if co_no is not None else "")
    if not co_no:
        return {}
    try:
        with ENGINE.connect() as conn:
            row = conn.execute(text(
                "SELECT id, co_no, project_name FROM projects "
                "WHERE co_no=:co AND project_type='parent'"
            ), {"co": co_no}).fetchone()
        if row:
            return {"id": row[0], "co_no": row[1] or "", "project_name": row[2] or ""}
    except Exception:
        pass
    return {}

def resolve_project_by_co_no(co_no) -> dict:
    """Look up ANY project (parent or child) strictly by co_no. Used by the Expenditures
    feature, which links to a project via co_no rather than the internal numeric id.
    Returns {} if not found or co_no is blank."""
    co_no = (str(co_no).strip() if co_no is not None else "")
    if not co_no:
        return {}
    try:
        with ENGINE.connect() as conn:
            row = conn.execute(text(
                "SELECT id, co_no, project_name FROM projects WHERE co_no=:co LIMIT 1"
            ), {"co": co_no}).fetchone()
        if row:
            return {"id": row[0], "co_no": row[1] or "", "project_name": row[2] or ""}
    except Exception as e:
        print(f"[resolve_project_by_co_no] {e}")
    return {}

def resolve_parent_by_label(label: str) -> dict:
    """Fallback for legacy/typed text: match a 'CO_NO-Project Name' label or a plain project
    name against known parents, so we can still recover the correct co_no."""
    label = (label or "").strip()
    if not label:
        return {}
    for opt in get_parent_project_options():
        if opt["label"].strip().lower() == label.lower() or opt["name"].strip().lower() == label.lower():
            return {"id": opt["id"], "co_no": opt["co_no"] or "", "project_name": opt["name"]}
    return {}

def get_child_project_names(parent_name: str = "") -> list[str]:
    try:
        with ENGINE.connect() as conn:
            if parent_name:
                rows = conn.execute(text(
                    "SELECT project_name FROM projects WHERE parent_project=:p ORDER BY project_name"
                ), {"p": parent_name}).fetchall()
            else:
                rows = conn.execute(text(
                    "SELECT project_name FROM projects WHERE project_type='child' ORDER BY project_name"
                )).fetchall()
        return [r[0] for r in rows]
    except Exception:
        return []


def add_ref_value(table: str, col: str, value: str) -> bool:
    try:
        with ENGINE.begin() as conn:
            conn.execute(text(
                f"INSERT INTO {table} ({col}) VALUES (:v) ON CONFLICT ({col}) DO NOTHING"
            ), {"v": value})
        return True
    except Exception:
        return False

def delete_ref_value(table: str, col: str, value: str) -> bool:
    try:
        with ENGINE.begin() as conn:
            conn.execute(text(f"UPDATE {table} SET active=FALSE WHERE {col}=:v"), {"v": value})
        return True
    except Exception:
        return False


# ── Per-LOB EBF multiplier settings ──────────────────────────
DEFAULT_EBF_MULT = 1.35

def get_ebf_mult(lob: str) -> float:
    """EBF multiplier saved for a specific Line of Business, falling back to
    the platform default (1.35) if none has been set yet."""
    if not lob or lob == "All Lines of Business" or lob == "All":
        return DEFAULT_EBF_MULT
    try:
        with ENGINE.connect() as conn:
            row = conn.execute(text(
                "SELECT ebf_mult FROM lob_ebf_settings WHERE lob=:l"
            ), {"l": lob}).fetchone()
        return float(row[0]) if row and row[0] is not None else DEFAULT_EBF_MULT
    except Exception as e:
        print(f"[get_ebf_mult] {e}")
        return DEFAULT_EBF_MULT

def get_all_ebf_mults() -> dict:
    """{lob: ebf_mult} for every LOB that has a custom value saved."""
    try:
        with ENGINE.connect() as conn:
            rows = conn.execute(text("SELECT lob, ebf_mult FROM lob_ebf_settings")).fetchall()
        return {r[0]: float(r[1]) for r in rows}
    except Exception as e:
        print(f"[get_all_ebf_mults] {e}")
        return {}

def save_ebf_mult(lob: str, mult: float) -> bool:
    if not lob or lob == "All Lines of Business" or lob == "All":
        return False
    try:
        mult = float(mult)
    except Exception:
        return False
    try:
        with ENGINE.begin() as conn:
            conn.execute(text("""
                INSERT INTO lob_ebf_settings (lob, ebf_mult, updated_at)
                VALUES (:l, :m, :t)
                ON CONFLICT (lob) DO UPDATE SET ebf_mult=:m, updated_at=:t
            """), {"l": lob, "m": mult, "t": datetime.now().isoformat()})
        return True
    except Exception as e:
        print(f"[save_ebf_mult] {e}")
        return False


# ═══════════════════════════════════════════════════════════════
#  CORE DB OPERATIONS
# ═══════════════════════════════════════════════════════════════

def _n(v, default=0.0):
    try:    return float(v or default)
    except: return default
def _rev_ov(p: dict) -> float:
    """Revised Order Value — falls back to Order Value + Variation if not set."""
    return _n(p.get("revised_order_value")) or (_n(p.get("order_value")) + _n(p.get("variation_amt")))

def _rows_to_dicts(conn, sql: str, params: dict = None) -> list[dict]:
    result = conn.execute(text(sql), params or {})
    keys   = list(result.keys())
    return [dict(zip(keys, row)) for row in result.fetchall()]


def _add_months(d: date, months: int) -> date:
  """Add months to a date, keeping end-of-month semantics."""
  if not isinstance(d, date):
    return d
  month = d.month - 1 + int(months)
  year = d.year + month // 12
  month = month % 12 + 1
  day = min(d.day, calendar.monthrange(year, month)[1])
  return date(year, month, day)


def save_project(fields: dict) -> int:
  num = [
    "order_value", "variation_amt", "revised_order_value", "baseline_eac", "baseline_margin",
    "rev_latest_eac", "rev_latest_margin", "physical_completion", "billed_amount",
    "collection_till_date", "itd_cost", "extension_months",
  ]
  for f in num:
    fields[f] = _n(fields.get(f))

  # auto-calc revised_order_value if not set
  if not fields["revised_order_value"]:
    fields["revised_order_value"] = fields["order_value"] + fields["variation_amt"]

  # Baseline Margin (%) = (1 - (Baseline EAC / Order Value)) x 100 (using Revised Order Value if present)
  # Rev. Latest Margin (%) = (1 - (Rev. Latest EAC / Order Value)) x 100 (using Revised Order Value if present)
  ov_for_margin = fields["revised_order_value"] or fields["order_value"]
  if ov_for_margin:
    fields["baseline_margin"]   = round((1 - (fields["baseline_eac"]   / ov_for_margin)) * 100, 2)
    fields["rev_latest_margin"] = round((1 - (fields["rev_latest_eac"] / ov_for_margin)) * 100, 2)
  else:
    fields["baseline_margin"]   = 0
    fields["rev_latest_margin"] = 0

  text_cols = ["co_no","project_name","business_line","project_manager","client_name",
         "order_number","order_issue_date","end_date","extended_end_date","hier_business_line",
         "parent_project","child_project","project_status",
         "last_billed_date","bill_submission_pending_date","folder_path"]
  for f in text_cols:
    fields.setdefault(f, "")

  # Only auto-calculate extended_end_date when the user hasn't manually typed one in.
  if fields["extension_months"] and not str(fields.get("extended_end_date") or "").strip():
    base_date = fields.get("end_date") or fields.get("order_issue_date") or ""
    if base_date:
      try:
        base_dt = datetime.fromisoformat(base_date).date()
      except Exception:
        try:
          base_dt = datetime.strptime(str(base_date)[:10], "%Y-%m-%d").date()
        except Exception:
          base_dt = None
      if base_dt:
        fields["extended_end_date"] = _add_months(base_dt, int(fields["extension_months"])) .isoformat()

  fields["created_at"] = datetime.now().isoformat()
  client = fields.get("client_name", "").strip()
  if client:
    add_ref_value("ref_clients", "name", client)

  parent_val      = fields.get("parent_project", "").strip()
  parent_id_raw   = fields.get("parent_id", "")
  submitted_type  = fields.get("project_type", "").strip()

  PLACEHOLDER_VALUES = {"", "select or type parent…", "none", "n/a"}
  has_parent_text = bool(parent_val) and parent_val.lower() not in PLACEHOLDER_VALUES

  # Resolve the parent by co_no first (co_no is unique per project) — matching on
  # typed/combined name text is what was causing sub-projects to save but not show up
  # under their parent.
  parent_row = {}
  if parent_id_raw or has_parent_text:
    parent_row = resolve_parent_by_co_no(parent_id_raw) or resolve_parent_by_label(parent_val)

  if parent_row:
    fields["project_type"]   = "child"
    fields["parent_id"]      = parent_row["co_no"]          # co_no, not the internal db id
    fields["parent_project"] = parent_row["project_name"]   # store the exact project name, not the "CO_NO-Name" label
  elif submitted_type == "child":
    fields["project_type"] = "child"
    fields["parent_id"]    = None
  else:
    fields["project_type"]   = "parent"
    fields["parent_project"] = ""
    fields["parent_id"]      = None

  new_id = None
  with ENGINE.begin() as conn:
    row = conn.execute(text("""
      INSERT INTO projects (
        co_no, project_name, business_line, project_manager, client_name,
      order_number, order_issue_date, end_date, extension_months, extended_end_date, order_value, variation_amt, revised_order_value,
        hier_business_line, parent_project, parent_id, child_project,
        baseline_eac, baseline_margin, rev_latest_eac, rev_latest_margin,
        project_status, physical_completion, billed_amount, last_billed_date,
        collection_till_date, itd_cost, bill_submission_pending_date,
        project_type, folder_path, created_at
      ) VALUES (
        :co_no, :project_name, :business_line, :project_manager, :client_name,
      :order_number, :order_issue_date, :end_date, :extension_months, :extended_end_date, :order_value, :variation_amt, :revised_order_value,
        :hier_business_line, :parent_project, :parent_id, :child_project,
        :baseline_eac, :baseline_margin, :rev_latest_eac, :rev_latest_margin,
        :project_status, :physical_completion, :billed_amount, :last_billed_date,
        :collection_till_date, :itd_cost, :bill_submission_pending_date,
        :project_type, :folder_path, :created_at
      ) RETURNING id"""), fields)
    new_id = row.fetchone()[0]

  child_name = fields.get("project_name", "").strip()
  if fields["project_type"] == "child" and child_name and new_id and (fields.get("parent_id") or parent_val):
    with ENGINE.begin() as conn2:
      if fields.get("parent_id"):
        conn2.execute(text("""
          UPDATE projects
          SET child_project = CASE
            WHEN child_project IS NULL OR child_project = ''
              THEN :child
            WHEN child_project NOT LIKE '%' || :child || '%'
              THEN child_project || ', ' || :child
            ELSE child_project
          END
          WHERE co_no = :pid
          AND   project_type = 'parent'
        """), {"child": child_name, "pid": fields["parent_id"]})
      else:
        conn2.execute(text("""
          UPDATE projects
          SET child_project = CASE
            WHEN child_project IS NULL OR child_project = ''
              THEN :child
            WHEN child_project NOT LIKE '%' || :child || '%'
              THEN child_project || ', ' || :child
            ELSE child_project
          END
          WHERE project_name = :parent
          AND   project_type = 'parent'
        """), {"child": child_name, "parent": parent_val})

  return new_id


def update_project(pid: int, fields: dict):
    num = ["order_value","variation_amt","revised_order_value","baseline_eac","baseline_margin",
           "rev_latest_eac","rev_latest_margin","physical_completion","billed_amount",
           "collection_till_date","itd_cost","extension_months"]
    for f in num:
        fields[f] = _n(fields.get(f))

    if not fields["revised_order_value"]:
        fields["revised_order_value"] = fields["order_value"] + fields["variation_amt"]

    # Baseline Margin (%) = (1 - (Baseline EAC / Order Value)) x 100 (using Revised Order Value if present)
    # Rev. Latest Margin (%) = (1 - (Rev. Latest EAC / Order Value)) x 100 (using Revised Order Value if present)
    ov_for_margin = fields["revised_order_value"] or fields["order_value"]
    if ov_for_margin:
        fields["baseline_margin"]   = round((1 - (fields["baseline_eac"]   / ov_for_margin)) * 100, 2)
        fields["rev_latest_margin"] = round((1 - (fields["rev_latest_eac"] / ov_for_margin)) * 100, 2)
    else:
        fields["baseline_margin"]   = 0
        fields["rev_latest_margin"] = 0

    text_cols = ["co_no","project_name","business_line","project_manager","client_name",
                 "order_number","order_issue_date","end_date","extended_end_date","hier_business_line",
                 "parent_project","child_project","project_status",
                 "last_billed_date","bill_submission_pending_date"]
    for f in text_cols:
        fields.setdefault(f, "")
    fields["folder_path"] = (fields.get("folder_path_check") or fields.get("folder_path") or "").strip()

    # Only auto-calculate extended_end_date when the user hasn't manually typed one in.
    if fields["extension_months"] and not str(fields.get("extended_end_date") or "").strip():
        base_date = fields.get("end_date") or fields.get("order_issue_date") or ""
        if base_date:
            try:
                base_dt = datetime.fromisoformat(base_date).date()
            except Exception:
                try:
                    base_dt = datetime.strptime(str(base_date)[:10], "%Y-%m-%d").date()
                except Exception:
                    base_dt = None
            if base_dt:
                fields["extended_end_date"] = _add_months(base_dt, int(fields["extension_months"])) .isoformat()

    fields["id"] = pid
    client = fields.get("client_name", "").strip()
    if client: add_ref_value("ref_clients", "name", client)
 
    parent_val      = fields.get("parent_project", "").strip()
    parent_id_raw   = fields.get("parent_id", "")
    submitted_type  = fields.get("project_type", "").strip()

    PLACEHOLDER_VALUES = {"", "select or type parent…", "none", "n/a"}
    has_parent_text = bool(parent_val) and parent_val.lower() not in PLACEHOLDER_VALUES

    # Resolve the parent by co_no first (co_no is unique per project) — matching on
    # typed/combined name text is what was causing sub-projects to save but not show up
    # under their parent.
    parent_row = {}
    if parent_id_raw or has_parent_text:
        parent_row = resolve_parent_by_co_no(parent_id_raw) or resolve_parent_by_label(parent_val)

    if parent_row:
        fields["project_type"]   = "child"
        fields["parent_id"]      = parent_row["co_no"]      # co_no, not the internal db id
        fields["parent_project"] = parent_row["project_name"]
    elif submitted_type == "child":
        fields["project_type"] = "child"
        fields["parent_id"]    = None
    else:
        fields["project_type"]   = "parent"
        fields["parent_project"] = ""
        fields["parent_id"]      = None
 
    with ENGINE.begin() as conn:
        conn.execute(text("""
            UPDATE projects SET
                co_no=:co_no,
                project_name=:project_name,
                project_type=:project_type,
                business_line=:business_line,
                project_manager=:project_manager,
                client_name=:client_name,
                order_number=:order_number,
                order_issue_date=:order_issue_date,
                end_date=:end_date,
                extension_months=:extension_months,
                extended_end_date=:extended_end_date,
                order_value=:order_value,
                variation_amt=:variation_amt,
                revised_order_value=:revised_order_value,
                hier_business_line=:hier_business_line,
                parent_project=:parent_project,
                parent_id=:parent_id,
                child_project=:child_project,
                project_status=:project_status,
                physical_completion=:physical_completion,
                billed_amount=:billed_amount,
                last_billed_date=:last_billed_date,
                collection_till_date=:collection_till_date,
                itd_cost=:itd_cost,
                bill_submission_pending_date=:bill_submission_pending_date,
                rev_latest_eac=:rev_latest_eac,
                rev_latest_margin=:rev_latest_margin,
                baseline_eac=:baseline_eac,
                baseline_margin=:baseline_margin,
                folder_path=:folder_path
            WHERE id=:id"""), fields)
 
    child_name = fields.get("project_name", "").strip()
    if fields["project_type"] == "child" and child_name and (fields.get("parent_id") or parent_val):
        with ENGINE.begin() as conn2:
            if fields.get("parent_id"):
                conn2.execute(text("""
                    UPDATE projects
                    SET child_project = CASE
                        WHEN child_project IS NULL OR child_project = '' THEN :child
                        WHEN child_project NOT LIKE '%' || :child || '%' THEN child_project || ', ' || :child
                        ELSE child_project
                    END
                    WHERE co_no = :pid AND project_type = 'parent'
                """), {"child": child_name, "pid": fields["parent_id"]})
            else:
                conn2.execute(text("""
                    UPDATE projects
                    SET child_project = CASE
                        WHEN child_project IS NULL OR child_project = '' THEN :child
                        WHEN child_project NOT LIKE '%' || :child || '%' THEN child_project || ', ' || :child
                        ELSE child_project
                    END
                    WHERE project_name = :parent AND project_type = 'parent'
                """), {"child": child_name, "parent": parent_val})

def delete_project(pid: int):
    with ENGINE.begin() as conn:
        row = conn.execute(text("SELECT co_no FROM projects WHERE id=:id"), {"id": pid}).fetchone()
        co_no = row[0] if row else None
        # expenditures/erp_updates/expense_entries are now linked by co_no, not project_id,
        # so ON DELETE CASCADE on project_id no longer cleans these up — do it explicitly.
        if co_no:
            for tbl in ("expense_entries", "erp_updates", "expenditures"):
                conn.execute(text(f"DELETE FROM {tbl} WHERE co_no=:c"), {"c": co_no})
        conn.execute(text("DELETE FROM projects WHERE id=:id"), {"id": pid})

def wipe_all_transactional_data():
    """Full wipe of every project-related table. Used when a new master
    Data Entry Excel is uploaded — the upload becomes a full replace, not
    a merge. Deleting from `projects` alone isn't enough if any FK is
    missing ON DELETE CASCADE, so every table is cleared explicitly."""
    with ENGINE.begin() as conn:
        for tbl in ["expense_entries", "erp_updates", "expenditures",
                    "billing_collection", "variations", "documents", "projects"]:
            try:
                conn.execute(text(f"DELETE FROM {tbl}"))
                print(f"[wipe_all_transactional_data] cleared {tbl}")
            except Exception as e:
                print(f"[wipe_all_transactional_data] {tbl}: {e}")
        for tbl in ["ref_lob", "ref_business_line"]:
            try:
                conn.execute(text(f"DELETE FROM {tbl}"))
                print(f"[wipe_all_transactional_data] cleared {tbl}")
            except Exception as e:
                print(f"[wipe_all_transactional_data] {tbl}: {e}")


def wipe_expenditure_data():
    """Wipe only the expenditure/cost tables. Used when a new Expenditure
    Excel is uploaded — projects themselves stay put, but ALL previous
    cost data is cleared before the new file's rows are matched in by
    project name / CO.No, so old figures never linger."""
    with ENGINE.begin() as conn:
        for tbl in ["expense_entries", "erp_updates", "expenditures"]:
            try:
                conn.execute(text(f"DELETE FROM {tbl}"))
                print(f"[wipe_expenditure_data] cleared {tbl}")
            except Exception as e:
                print(f"[wipe_expenditure_data] {tbl}: {e}")
        try:
            conn.execute(text("UPDATE projects SET itd_cost = 0"))
        except Exception as e:
            print(f"[wipe_expenditure_data] reset itd_cost: {e}")


def fetch_all_projects(lob_filter: str = "") -> list[dict]:
        try:
                with ENGINE.connect() as conn:
                    if lob_filter and lob_filter != "All":
                        rows = _rows_to_dicts(conn,
                            "SELECT * FROM projects WHERE hier_business_line=:l OR business_line=:l ORDER BY id DESC",
                            {"l": lob_filter})
                    else:
                        rows = _rows_to_dicts(conn, "SELECT * FROM projects ORDER BY id DESC")

                role = (getattr(_req_ctx, "role", "") or "").lower()
                if role == "pm":
                    pm_name = (getattr(_req_ctx, "pm_name", "") or "").strip().lower()
                    rows = [r for r in rows if (r.get("project_manager") or "").strip().lower() == pm_name]

                # Global "Active Projects" toggle (topbar button) -- when on, every
                # page that lists/aggregates projects via this function sees only
                # Active ones, until the user switches it back off.
                if getattr(_req_ctx, "active_only", False):
                    rows = [r for r in rows if (r.get("project_status") or "").strip().lower() == "active"]

                return rows
        except Exception as e:
            print(f"[fetch_all_projects] {e}")
        return []

def fetch_project_by_id(pid: int) -> dict:
    try:
        with ENGINE.connect() as conn:
            result = conn.execute(text("SELECT * FROM projects WHERE id=:id"), {"id": pid})
            keys = list(result.keys())
            row = result.fetchone()
            if row:
                return dict(zip(keys, row))
    except Exception as e:
        print(f"[fetch_project_by_id] {e}")
    return {}


def fetch_parent_with_children_totals(parent_name: str, exclude_id: int = 0) -> dict:
  try:
    with ENGINE.connect() as conn:
      rows = conn.execute(text(
        "SELECT * FROM projects WHERE project_name=:n AND project_type='parent' LIMIT 1"
      ), {"n": parent_name}).fetchall()
      if not rows:
        return {}
      par = dict(zip(rows[0].keys(), rows[0])) if hasattr(rows[0], '__iter__') else dict(rows[0])

      if exclude_id:
        children = _rows_to_dicts(conn,
          "SELECT * FROM projects WHERE parent_project=:n AND id != :ex",
          {"n": parent_name, "ex": exclude_id})
      else:
        children = _rows_to_dicts(conn,
          "SELECT * FROM projects WHERE parent_project=:n",
          {"n": parent_name})

    def s(k):
      return sum(_n(c.get(k)) for c in children)

    par_ov = _rev_ov(par)
    par_eac = _n(par.get("rev_latest_eac") or par.get("baseline_eac"))
    par_bil = _n(par.get("billed_amount"))
    par_col = _n(par.get("collection_till_date"))
    par_itd = _n(par.get("itd_cost"))
    par_base_eac = _n(par.get("baseline_eac"))

    used_ov = sum(_rev_ov(c) for c in children)
    used_eac = s("rev_latest_eac") or s("baseline_eac")
    used_bil = s("billed_amount")
    used_col = s("collection_till_date")
    used_itd = s("itd_cost")
    used_base = s("baseline_eac")

    return {
      "parent_id": par.get("id"),
      "parent_name": par.get("project_name"),

      "parent_co_no": par.get("co_no") or "",
      "parent_business_line": par.get("hier_business_line") or par.get("business_line") or "",
      "parent_project_manager": par.get("project_manager") or "",
      "parent_client_name": par.get("client_name") or "",
      "parent_order_number": par.get("order_number") or "",
      "parent_order_issue_date": par.get("order_issue_date") or "",
      "parent_status": par.get("project_status") or "Active",
      "parent_physical_completion": _n(par.get("physical_completion")),

      "parent_baseline_eac": par_base_eac,
      "parent_baseline_margin": _n(par.get("baseline_margin")),
      "parent_rev_latest_margin": _n(par.get("rev_latest_margin")),

      "parent_order_value": par_ov,
      "parent_eac": par_eac,
      "parent_billed": par_bil,
      "parent_collection": par_col,
      "parent_itd_cost": par_itd,
      "child_count": len(children),

      "used_order_value": used_ov,
      "used_eac": used_eac,
      "used_billed": used_bil,
      "used_collection": used_col,
      "used_itd_cost": used_itd,
      "used_baseline_eac": used_base,

      "rem_order_value": par_ov - used_ov,
      "rem_eac": par_eac - used_eac,
      "rem_billed": par_bil - used_bil,
      "rem_collection": par_col - used_col,
      "rem_itd_cost": par_itd - used_itd,
      "rem_baseline_eac": par_base_eac - used_base,
    }
  except Exception as e:
    print(f"[fetch_parent_with_children_totals] {e}")
    return {}
 
def fetch_projects_for_select():
    try:
        with ENGINE.connect() as conn:
            rows = conn.execute(text(
                "SELECT id, co_no, project_name, client_name, order_number, order_value, "
                "business_line, variation_amt, revised_order_value, hier_business_line, "
                "project_manager "
                "FROM projects ORDER BY id DESC"
            )).fetchall()
        return rows
    except Exception:
        return []


# ── Documents ──────────────────────────────────────────────

def link_document(fields: dict):
    fields["created_at"] = datetime.now().isoformat()
    with ENGINE.begin() as conn:
        conn.execute(text("""
            INSERT INTO documents (project_id,doc_name,doc_category,doc_link,created_at)
            VALUES (:project_id,:doc_name,:doc_category,:doc_link,:created_at)
        """), fields)

def delete_document(doc_id: int):
    with ENGINE.begin() as conn:
        conn.execute(text("DELETE FROM documents WHERE id=:id"), {"id": doc_id})

def fetch_documents(project_id: int = None) -> list:
    try:
        with ENGINE.connect() as conn:
            if project_id:
                rows = conn.execute(text(
                    "SELECT id,doc_name,doc_category,doc_link FROM documents WHERE project_id=:p"
                ), {"p": project_id}).fetchall()
            else:
                rows = conn.execute(text(
                    "SELECT id,doc_name,doc_category,doc_link FROM documents ORDER BY id DESC"
                )).fetchall()
        return rows
    except Exception:
        return []


# ── Expenditures ───────────────────────────────────────────
def save_expenditure(fields: dict) -> int:
    # co_no (e.g. "620_7") is now the linking key, not the internal numeric project_id.
    co_no = (str(fields.get("co_no") or fields.get("project_id") or "")).strip()
    fields["co_no"] = co_no

    if not co_no:
        print("[save_expenditure] Rejected: no co_no supplied")
        return 0

    proj = resolve_project_by_co_no(co_no)
    pid_int = proj.get("id")
    if not pid_int:
        print(f"[save_expenditure] Rejected: no project found for co_no={co_no!r}")
        return 0

    # map of erp_<field> -> category name used in erp_updates / ref_expense_category
    legacy_map = {"erp_cogs":"COGS","erp_djcs":"DJCS","erp_equip":"Equipment",
                  "erp_manpower":"Manpower","erp_vehicles":"Vehicles","erp_others":"Others"}

    fixed_num = ("erp_cogs","erp_djcs","erp_equip","erp_manpower","erp_vehicles","erp_others",
                 "itd_cogs","itd_djcs","itd_equip","itd_manpower","itd_vehicles","itd_others")
    for f in fixed_num:
        fields[f] = _n(fields.get(f))

    fields.setdefault("created_at", datetime.now().isoformat())
    now_iso = datetime.now().isoformat()

    with ENGINE.begin() as conn:
        # 1) Fetch *current* cumulative totals per category from erp_updates history
        existing_totals = dict(conn.execute(text(
            "SELECT category, SUM(amount) FROM erp_updates WHERE co_no=:co GROUP BY category"
        ), {"co": co_no}).fetchall())

        # 2) For every ERP field the user typed a non-zero "this update" amount into,
        #    insert a new history row (only if > 0 — don't log no-op updates)
        for erp_field, category in legacy_map.items():
            delta = _n(fields.get(erp_field))
            if delta:
                conn.execute(text("""
                    INSERT INTO erp_updates (project_id, co_no, category, amount, created_by, created_at)
                    VALUES (:pid, :co, :cat, :amt, :by, :ts)
                """), {"pid": pid_int, "co": co_no, "cat": category, "amt": delta,
                       "by": fields.get("created_by", "system"), "ts": now_iso})

        # 3) Recompute fresh cumulative totals (existing + this update)
        new_totals = {}
        for erp_field, category in legacy_map.items():
            prev = _n(existing_totals.get(category, 0))
            delta = _n(fields.get(erp_field))
            new_totals[erp_field] = prev + delta

        # 4) Store the *cumulative* values into expenditures.erp_* (snapshot row),
        #    not the raw delta the user typed
        snapshot_fields = {**fields, **new_totals, "project_id": pid_int}

        existing = conn.execute(text(
            "SELECT id FROM expenditures WHERE co_no=:co ORDER BY id DESC LIMIT 1"
        ), {"co": co_no}).fetchone()

        if existing:
            exp_id = existing[0]
            conn.execute(text("""
                UPDATE expenditures SET
                    co_no=:co_no, lob=:lob, data_as_of=:data_as_of, erp_update_date=:erp_update_date,
                    itd_cogs=:itd_cogs, itd_djcs=:itd_djcs, itd_equip=:itd_equip,
                    itd_manpower=:itd_manpower, itd_vehicles=:itd_vehicles, itd_others=:itd_others,
                    erp_cogs=:erp_cogs, erp_djcs=:erp_djcs, erp_equip=:erp_equip,
                    erp_manpower=:erp_manpower, erp_vehicles=:erp_vehicles, erp_others=:erp_others
                WHERE id=:id
            """), {**snapshot_fields, "id": exp_id})
        else:
            r = conn.execute(text("""
                INSERT INTO expenditures (
                    project_id, co_no, lob, data_as_of, erp_update_date,
                    itd_cogs, itd_djcs, itd_equip, itd_manpower, itd_vehicles, itd_others,
                    erp_cogs, erp_djcs, erp_equip, erp_manpower, erp_vehicles, erp_others,
                    created_at
                ) VALUES (
                    :project_id, :co_no, :lob, :data_as_of, :erp_update_date,
                    :itd_cogs, :itd_djcs, :itd_equip, :itd_manpower, :itd_vehicles, :itd_others,
                    :erp_cogs, :erp_djcs, :erp_equip, :erp_manpower, :erp_vehicles, :erp_others,
                    :created_at
                ) RETURNING id"""), snapshot_fields)
            exp_id = r.fetchone()[0]

        total_erp = sum(new_totals.values())
        cur = conn.execute(text("SELECT itd_cost FROM projects WHERE id=:id"), {"id": pid_int}).fetchone()
        existing_itd = float(cur[0] or 0.0) if cur else 0.0
        final_itd = total_erp if total_erp >= existing_itd else existing_itd
        conn.execute(text(
            "UPDATE projects SET itd_cost=:v WHERE id=:id"
        ), {"v": final_itd, "id": pid_int})

    return exp_id


def fetch_expenditure_for_co_no(co_no: str) -> dict:
    try:
        with ENGINE.connect() as conn:
            rows = _rows_to_dicts(conn,
                "SELECT * FROM expenditures WHERE co_no=:c ORDER BY id DESC LIMIT 1",
                {"c": co_no})
        return rows[0] if rows else {}
    except Exception:
        return {}

# Backward-compat alias in case any old caller still passes the numeric project_id
def fetch_expenditure_for_project(project_id) -> dict:
    proj = resolve_project_by_co_no(project_id) if isinstance(project_id, str) and not project_id.isdigit() else None
    if proj:
        return fetch_expenditure_for_co_no(proj["co_no"])
    try:
        with ENGINE.connect() as conn:
            row = conn.execute(text("SELECT co_no FROM projects WHERE id=:id"), {"id": project_id}).fetchone()
        return fetch_expenditure_for_co_no(row[0]) if row else {}
    except Exception:
        return {}

def fetch_erp_totals(co_no: str) -> dict:
    """Returns {category: cumulative_total} for a project, summed from erp_updates.
    Safe to call even if the table doesn't exist yet or has no rows."""
    if not co_no:
        return {}
    try:
        with ENGINE.connect() as conn:
            rows = conn.execute(text(
                "SELECT category, SUM(amount) FROM erp_updates WHERE co_no=:co GROUP BY category"
            ), {"co": co_no}).fetchall()
        return {cat: (amt or 0) for cat, amt in rows}
    except Exception as e:
        print(f"[fetch_erp_totals] {e}")
        return {}

def fetch_erp_update_entries(co_no: str) -> list[dict]:
    """Returns each individual ERP reconciliation update row (id, category, amount,
    created_by, created_at) for a project — the raw log behind fetch_erp_totals()'s
    cumulative sums. Used to show "how much added and when" per head."""
    if not co_no:
        return []
    try:
        with ENGINE.connect() as conn:
            return _rows_to_dicts(conn,
                "SELECT * FROM erp_updates WHERE co_no=:co ORDER BY id DESC",
                {"co": co_no})
    except Exception as e:
        print(f"[fetch_erp_update_entries] {e}")
        return []

def save_expense_entry(fields: dict) -> int:
    for f in ("quantity","unit_cost"): fields[f] = _n(fields.get(f))
    fields["line_total"] = fields["quantity"] * fields["unit_cost"]
    fields["created_at"] = datetime.now().isoformat()

    cat = fields.get("category","").strip()
    if cat: add_ref_value("ref_expense_category","name",cat)

    co_no = (str(fields.get("co_no") or fields.get("project_id") or "")).strip()
    fields["co_no"] = co_no

    proj = resolve_project_by_co_no(co_no) if co_no else {}
    fields["project_id"] = proj.get("id")

    exp_id = int(fields.get("expenditure_id") or 0)
    if exp_id <= 0 and co_no:
        existing = fetch_expenditure_for_co_no(co_no)
        if existing:
            exp_id = existing["id"]
        else:
            stub = {
                "co_no": co_no, "lob": "", "data_as_of": date.today().isoformat(),
                "erp_update_date": date.today().isoformat(),
                "itd_cogs":0,"itd_djcs":0,"itd_equip":0,"itd_manpower":0,"itd_vehicles":0,"itd_others":0,
                "erp_cogs":0,"erp_djcs":0,"erp_equip":0,"erp_manpower":0,"erp_vehicles":0,"erp_others":0,
            }
            exp_id = save_expenditure(stub)

    fields["expenditure_id"] = exp_id

    with ENGINE.begin() as conn:
        r = conn.execute(text("""
            INSERT INTO expense_entries
                (expenditure_id, project_id, co_no, category, description,
                 quantity, unit_cost, line_total, created_at)
            VALUES
                (:expenditure_id, :project_id, :co_no, :category, :description,
                 :quantity, :unit_cost, :line_total, :created_at)
            RETURNING id"""), fields)
        return r.fetchone()[0]


def delete_expense_entry(entry_id: int):
    with ENGINE.begin() as conn:
        conn.execute(text("DELETE FROM expense_entries WHERE id=:id"), {"id": entry_id})


def fetch_expense_entries(co_no=None) -> list[dict]:
    try:
        with ENGINE.connect() as conn:
            if co_no:
                return _rows_to_dicts(conn,
                    "SELECT * FROM expense_entries WHERE co_no=:c ORDER BY id DESC",
                    {"c": co_no})
            else:
                return _rows_to_dicts(conn, "SELECT * FROM expense_entries ORDER BY id DESC")
    except Exception:
        return []


# ── Variations ─────────────────────────────────────────────

def save_variation(fields: dict) -> int:
    for f in ("original_value","variation_amt","revised_value"): fields[f] = _n(fields.get(f))
    fields["created_at"] = datetime.now().isoformat()
    with ENGINE.begin() as conn:
        r = conn.execute(text("""
            INSERT INTO variations
              (project_id, child_co_no, child_name, parent_name,
               original_value, variation_amt, revised_value, variation_date, status, created_at)
            VALUES
              (:project_id, :child_co_no, :child_name, :parent_name,
               :original_value, :variation_amt, :revised_value, :variation_date, :status, :created_at)
            RETURNING id"""), {**fields, "status": fields.get("status","Pending")})
        return r.fetchone()[0]


def delete_variation(var_id: int):
    with ENGINE.begin() as conn:
        conn.execute(text("DELETE FROM variations WHERE id=:id"), {"id": var_id})


def fetch_all_variations() -> list[dict]:
    try:
        with ENGINE.connect() as conn:
            return _rows_to_dicts(conn, "SELECT * FROM variations ORDER BY id DESC")
    except Exception:
        return []
def save_billing_entry(fields: dict) -> int:
    pid_raw = fields.get("project_id", 0)
    pid = int(pid_raw) if str(pid_raw).strip().isdigit() else 0
    if pid <= 0:
        return 0
    fields["project_id"]     = pid
    fields["billed_amt"]     = _n(fields.get("billed_amt"))
    fields["collection_amt"] = _n(fields.get("collection_amt"))
    fields.setdefault("ra_no", "")
    fields.setdefault("billing_date", "")
    fields.setdefault("collection_date", "")
    fields["created_at"] = datetime.now().isoformat()
    with ENGINE.begin() as conn:
        r = conn.execute(text("""
            INSERT INTO billing_collection
                (project_id, ra_no, billing_date, billed_amt, collection_date, collection_amt, created_at)
            VALUES
                (:project_id, :ra_no, :billing_date, :billed_amt, :collection_date, :collection_amt, :created_at)
            RETURNING id"""), fields)
        new_id = r.fetchone()[0]
    _sync_project_billing_totals(pid)
    return new_id
 
 
def update_billing_collection(entry_id: int, collection_date: str, collection_amt) -> bool:
    try:
        with ENGINE.begin() as conn:
            row = conn.execute(text(
                "SELECT project_id FROM billing_collection WHERE id=:id"
            ), {"id": entry_id}).fetchone()
            if not row:
                return False
            conn.execute(text("""
                UPDATE billing_collection
                SET collection_date=:cd, collection_amt=:ca
                WHERE id=:id
            """), {"cd": collection_date, "ca": _n(collection_amt), "id": entry_id})
        _sync_project_billing_totals(row[0])
        return True
    except Exception:
        return False
 
 
def delete_billing_entry(entry_id: int):
    row = None
    with ENGINE.begin() as conn:
        row = conn.execute(text(
            "SELECT project_id FROM billing_collection WHERE id=:id"
        ), {"id": entry_id}).fetchone()
        conn.execute(text("DELETE FROM billing_collection WHERE id=:id"), {"id": entry_id})
    if row:
        _sync_project_billing_totals(row[0])
 
 
def fetch_billing_entries(project_id: int) -> list:
    if not project_id:
        return []
    try:
        with ENGINE.connect() as conn:
            return _rows_to_dicts(conn,
                "SELECT * FROM billing_collection WHERE project_id=:p ORDER BY id DESC",
                {"p": project_id})
    except Exception:
        return []
def fetch_billing_entries_filtered(project_id=None, lob=None, pending_only=False) -> list:
    """Returns billing_collection rows joined with project name/CO.No,
    optionally narrowed by project_id and/or LOB, optionally pending-only."""
    try:
        with ENGINE.connect() as conn:
            sql = """
                SELECT bc.*, p.project_name, p.co_no, p.hier_business_line, p.business_line
                FROM billing_collection bc
                JOIN projects p ON p.id = bc.project_id
                WHERE 1=1
            """
            params = {}
            if project_id:
                sql += " AND bc.project_id = :pid"
                params["pid"] = project_id
            if lob:
                sql += " AND (p.hier_business_line = :lob OR p.business_line = :lob)"
                params["lob"] = lob
            if pending_only:
                sql += " AND (bc.collection_date IS NULL OR bc.collection_date = '')"
            sql += " ORDER BY bc.id DESC"
            return _rows_to_dicts(conn, sql, params)
    except Exception as e:
        print(f"[fetch_billing_entries_filtered] {e}")
        return [] 
 
def fetch_billing_summary_for_pm(pm_name: str) -> list[dict]:
    """Returns one row per project assigned to the given PM, with that
    project's total billed and total collection amounts summed across all
    its billing_collection entries (0 for projects with no entries yet)."""
    try:
        with ENGINE.connect() as conn:
            return _rows_to_dicts(conn, """
                SELECT p.id AS project_id, p.co_no, p.project_name,
                       COALESCE(SUM(bc.billed_amt),0)     AS total_billed,
                       COALESCE(SUM(bc.collection_amt),0) AS total_collection,
                       COUNT(bc.id)                        AS entry_count
                FROM projects p
                LEFT JOIN billing_collection bc ON bc.project_id = p.id
                WHERE p.project_manager = :pm
                GROUP BY p.id, p.co_no, p.project_name
                ORDER BY p.project_name
            """, {"pm": pm_name})
    except Exception as e:
        print(f"[fetch_billing_summary_for_pm] {e}")
        return []


def _sync_project_billing_totals(project_id: int):
    """Keeps projects.billed_amount / collection_till_date in sync with the
    billing_collection ledger, so Financials/KPI/Dashboard stay accurate
    automatically whenever a billing or collection entry changes."""
    try:
        with ENGINE.begin() as conn:
            row = conn.execute(text("""
                SELECT COALESCE(SUM(billed_amt),0), COALESCE(SUM(collection_amt),0)
                FROM billing_collection WHERE project_id=:p
            """), {"p": project_id}).fetchone()
            billed_total, collection_total = row[0], row[1]
            conn.execute(text("""
                UPDATE projects SET billed_amount=:b, collection_till_date=:c
                WHERE id=:id
            """), {"b": billed_total, "c": collection_total, "id": project_id})
    except Exception as e:
        print(f"[_sync_project_billing_totals] {e}")


def _sync_project_itd_cost(project_id: int):
    """Keeps projects.itd_cost in sync with the erp_updates ledger, so the
    Financial Performance table (and every other view that reads
    projects.itd_cost) reflects newly-added/imported expenditure data
    automatically — without requiring a manual 'Update Data' click on the
    Expenditures page. Preserves any higher authoritative ITD cost (e.g. from
    Projects sheet Column U)."""
    try:
        with ENGINE.begin() as conn:
            row = conn.execute(text("""
                SELECT co_no, itd_cost FROM projects WHERE id=:id
            """), {"id": project_id}).fetchone()
            if not row or not row[0]:
                return
            co_no = row[0]
            existing_itd = float(row[1] or 0.0)
            total = conn.execute(text("""
                SELECT COALESCE(SUM(amount),0) FROM erp_updates WHERE co_no=:co
            """), {"co": co_no}).fetchone()[0]
            total = float(total or 0.0)
            final_itd = total if total >= existing_itd else (existing_itd if existing_itd > 0 else total)
            conn.execute(text("""
                UPDATE projects SET itd_cost=:v WHERE id=:id
            """), {"v": final_itd, "id": project_id})
    except Exception as e:
        print(f"[_sync_project_itd_cost] {e}")



def fetch_summary_stats() -> dict:
    projects   = fetch_all_projects()
    variations = fetch_all_variations()
    parents    = [p for p in projects if not (p.get("parent_project") or "").strip()]
    children   = [p for p in projects if (p.get("parent_project") or "").strip()]
    net_var    = sum(_n(v.get("variation_amt")) for v in variations)
    pending    = sum(1 for v in variations if (v.get("status") or "Pending") == "Pending")
    return {
        "total_parents":  len(parents)  or 0,
        "total_children": len(children) or 0,
        "net_variation":  net_var,
        "pending":        pending,
    }
import html as _htmlmod

def search_projects(query: str) -> list[dict]:
    if not query or not query.strip():
        return []
    like = f"%{query.strip()}%"
    try:
        with ENGINE.connect() as conn:
            return _rows_to_dicts(conn, """
                SELECT * FROM projects
                WHERE project_name      ILIKE :q
                   OR co_no             ILIKE :q
                   OR client_name       ILIKE :q
                   OR project_manager   ILIKE :q
                   OR order_number      ILIKE :q
                   OR business_line     ILIKE :q
                   OR hier_business_line ILIKE :q
                ORDER BY id DESC
            """, {"q": like})
    except Exception as e:
        print(f"[search_projects] {e}")
        return []


def search_results_page(query: str, username="Admin") -> str:
    safe_q  = _htmlmod.escape(query or "")
    results = search_projects(query)

    rows_html = ""
    for r in results:
        pid    = r.get("id")
        co     = _htmlmod.escape(r.get("co_no") or "—")
        name   = _htmlmod.escape(r.get("project_name") or "—")
        pm     = _htmlmod.escape(r.get("project_manager") or "—")
        client = _htmlmod.escape(r.get("client_name") or "—")
        status = _htmlmod.escape(r.get("project_status") or "—")
        lob    = _htmlmod.escape(r.get("hier_business_line") or r.get("business_line") or "—")
        rows_html += f"""
        <tr style="border-bottom:1px solid {C['outline']}">
          <td style="padding:12px 14px;font-weight:600;color:{C['dark']}">{co}</td>
          <td style="padding:12px 14px;font-weight:600">{name}</td>
          <td style="padding:12px 14px">{lob}</td>
          <td style="padding:12px 14px">{pm}</td>
          <td style="padding:12px 14px">{client}</td>
          <td style="padding:12px 14px">{status}</td>
          <td style="padding:12px 14px;text-align:right;white-space:nowrap">
            <a href="/edit_project?id={pid}" class="btn-outline" style="margin-right:6px">✏ Edit</a>
            <a href="/expenditures?project_id={pid}" class="btn-outline" style="margin-right:6px">Expenditures</a>
            <a href="/billing_collection?project_id={pid}" class="btn-outline">Billing</a>
          </td>
        </tr>"""

    if not rows_html:
        rows_html = (f'<tr><td colspan="7" style="padding:28px;text-align:center;color:#aaa">'
                     f'No projects matched &ldquo;{safe_q}&rdquo;.</td></tr>')

    html  = _html_head("Search Results") + "<body>"
    html += _sidebar("") + _topbar("SEARCH RESULTS", username, search_value=query)
    html += f"""
<div class="main"><div class="canvas">
  <h2 style="font-size:22px;font-weight:800;color:{C['dark']};margin-bottom:4px">Search Results</h2>
  <p style="font-size:13px;color:#6b7280;margin-bottom:20px">
    {len(results)} project(s) matched &ldquo;{safe_q}&rdquo;
  </p>
  <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;overflow:hidden">
    <div style="overflow-x:auto">
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead>
          <tr style="background:{C['surface_low']};border-bottom:1px solid {C['outline']}">
            <th style="text-align:left;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">CO.No</th>
            <th style="text-align:left;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Project Name</th>
            <th style="text-align:left;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">LOB</th>
            <th style="text-align:left;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">PM</th>
            <th style="text-align:left;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Client</th>
            <th style="text-align:left;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Status</th>
            <th style="text-align:right;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Actions</th>
          </tr>
        </thead>
        <tbody>{rows_html}</tbody>
      </table>
    </div>
  </div>
</div></div>
</body></html>"""
    return html

# ═══════════════════════════════════════════════════════════════
#  SHARED HTML HELPERS
# ═══════════════════════════════════════════════════════════════

NAV_ITEMS = [
    ("dashboard",   "💻",  "Dashboard"),
    ("chat",        "💬",  "Chat"),
    ("data_entry",  "🖊️",  "Data Entry"),
    ("expenditures","💳", "Expenditures"),
    ("financials",  "📊", "Financials Performance"),
    ("variations",  "📈", "Parent Project Details"),
    ("kpi",         "🎯", "KPI Monitoring"),
    ("eac_status",  "📐", "EAC Status Report"),
    ("schedule",    "📅", "Schedule Report"),
    ("billing_collection", "🧾", "Billing & Collection"),
]

# Weekly Payment Booking / Payment Approvals now live in their own separate
# portal (see PAYMENTS_NAV_ITEMS / _payments_sidebar / payments_dashboard),
# reached via the portal switch on the login page — they no longer appear
# in the main system sidebar.
PAYMENTS_NAV_ITEMS = [
    ("payments_dashboard", "🏦", "Payments Dashboard"),
    ("weekly_booking",     "🗓️", "Weekly Payment Booking"),
    ("payment_approvals",  "✅", "Payment Approvals"),
]

# Keys from NAV_ITEMS that a "director" role is allowed to see/access
# (CFO, CEO, and other directors all log in under the "director" role, so this
# also governs what they see — "billing_collection" is included so they can view
# Billing & Collection Details.)
DIRECTOR_NAV_KEYS = {"dashboard", "variations", "reports", "kpi", "eac_status", "financials", "schedule", "billing_collection"}

# Full URL paths (including export endpoints) a director may load directly
DIRECTOR_ALLOWED_PATHS = {
    "/dashboard", "/", "/index",
    "/variations", "/reports", "/kpi", "/eac_status", "/financials", "/chat", "/chat_sessions", "/chat_history",
    "/schedule",
    "/export_excel", "/export_excel/full",
    "/export_excel/variations", "/export_pdf/variations",
    "/export_excel/reports",
    "/export_excel/kpi",
    "/export_excel/eac_status",
    "/export_excel/financials", "/export_pdf/financials",
    "/export_excel/schedule",
    "/billing_collection", "/billing_collection_retrieve", "/export_excel/billing",
    "/payments_dashboard",
    "/payment_approvals", "/export_excel/payment_bookings", "/download_stamped_doc",
    "/toggle_active_filter",
}

# POST paths a director may submit (all other POSTs redirect to /variations, as before)
DIRECTOR_ALLOWED_POST_PATHS = {"/director_decide", "/financials_filter", "/save_ebf_lob", "/billing_collection_retrieve", "/delete_chat_session", "/clear_chat_history"}

# Keys from NAV_ITEMS that a "pm"/"head" (Project Manager / Head of Dept) role is NOT allowed to see/access
# NOTE: "billing_collection" is intentionally NOT hidden — a PM login is allowed to see the
# Billing & Collection Details tab, but it is scoped server-side (see billing_collection())
# so a PM only ever sees their own billing entries, never anyone else's.
PM_HIDDEN_NAV_KEYS = {"data_entry", "expenditures"}

# Full URL paths (including export endpoints) a pm/head may NOT load directly
PM_BLOCKED_PATHS = {
    "/data_entry", "/expenditures",
    "/export_pdf/data_entry",
    "/export_excel/expenditures",
    "/expenditures_retrieve",
    "/payment_approvals", "/procurement_decide", "/accounts_decide", "/director_decide",
}

# Keys from NAV_ITEMS an "accounts" role may see in the main sidebar (payments
# portal pages use PAYMENTS_NAV_ITEMS / _payments_sidebar instead)
ACCOUNTS_NAV_KEYS = {"chat"}
# Full URL paths an accounts login may load directly
ACCOUNTS_ALLOWED_PATHS = {
    "/payments_dashboard",
    "/payment_approvals", "/export_excel/payment_bookings", "/download_stamped_doc", "/chat", "/chat_sessions", "/chat_history",
    "/toggle_active_filter",
}
ACCOUNTS_ALLOWED_POST_PATHS = {"/accounts_decide", "/delete_chat_session", "/clear_chat_history"}

# "procurement" gets both Screen 1 (data entry, per spec) and Screen 2 (approvals),
# both now on the separate payments portal — see PAYMENTS_NAV_ITEMS.
PROCUREMENT_NAV_KEYS = {"chat"}
PROCUREMENT_ALLOWED_PATHS = {
    "/payments_dashboard",
    "/weekly_booking", "/payment_approvals",
    "/export_excel/payment_bookings", "/download_stamped_doc", "/chat", "/chat_sessions", "/chat_history",
    "/toggle_active_filter",
}
PROCUREMENT_ALLOWED_POST_PATHS = {"/save_booking", "/procurement_decide", "/delete_chat_session", "/clear_chat_history"}

SHARED_CSS = f"""
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',Arial,sans-serif;background:{C['bg']};color:{C['on_surface']};display:flex;min-height:100vh}}
.sidebar{{width:200px;min-height:100vh;background:#fff;border-right:1px solid {C['outline']};display:flex;flex-direction:column;padding:20px 0 16px;position:fixed;top:0;left:0;z-index:200}}
.brand{{padding:0 16px 20px;font-size:13px;font-weight:700;line-height:1.4;color:{C['dark']};border-bottom:1px solid #eee;margin-bottom:8px}}
.nav-item{{display:flex;align-items:center;gap:9px;padding:9px 16px;text-decoration:none;font-size:13px;color:#4a5568;border-left:3px solid transparent;transition:background .15s;white-space:normal;line-height:1.3}}
.nav-item:hover{{background:{C['surface_high']}}}
.nav-item.active{{background:#f0f3fa;color:{C['dark']};font-weight:700;border-left-color:{C['dark']}}}
.nav-icon{{font-size:14px;width:18px;text-align:center;flex-shrink:0}}
.sidebar-btn{{margin:auto 14px 0;background:{C['dark']};color:#fff;border:none;border-radius:6px;padding:9px 0;font-size:12px;font-weight:700;cursor:pointer;width:calc(100% - 28px)}}
.topbar{{position:fixed;top:0;left:200px;right:0;height:52px;background:#fff;border-bottom:1px solid {C['outline']};display:flex;align-items:center;justify-content:space-between;padding:0 24px;z-index:199}}
.topbar-title{{font-size:12px;font-weight:700;color:#6b7280;letter-spacing:.04em;text-transform:uppercase}}
.topbar-right{{display:flex;align-items:center;gap:14px}}
.search-box{{background:{C['surface_high']};border:1px solid {C['outline']};border-radius:6px;padding:6px 12px;font-size:12px;color:#9ca3af;width:180px;outline:none}}
.topbar-nav a{{font-size:12px;font-weight:600;color:{C['on_var']};text-decoration:none;margin-left:18px}}
.topbar-nav a:hover{{color:{C['dark']}}}
.proj-sel-btn{{background:{C['dark']};color:#fff;border:none;border-radius:4px;padding:6px 12px;font-size:12px;font-weight:700;cursor:pointer}}
.tb-portal-switch{{display:flex;gap:2px;background:{C['surface_high']};border:1px solid {C['outline']};border-radius:20px;padding:3px}}
.tb-portal-btn{{border:none;background:transparent;color:{C['on_var']};font-size:11.5px;font-weight:700;
  letter-spacing:.01em;padding:6px 13px;border-radius:16px;cursor:pointer;text-decoration:none;
  transition:background .15s,color .15s;white-space:nowrap;display:inline-block}}
.tb-portal-btn.active{{background:{C['dark']};color:#fff}}
.tb-portal-btn:not(.active):hover{{color:{C['dark']}}}
.avatar{{background:{C['dark']};color:#fff;border-radius:50%;width:28px;height:28px;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700}}
.avatar-link{{text-decoration:none;display:flex;align-items:center;gap:8px;color:{C['on_surface']}}}
.main{{margin-left:200px;padding-top:52px;flex:1;min-width:0;min-height:100vh;background:{C['bg']}}}
.canvas{{max-width:1280px;margin:0 auto;padding:24px}}
.btn-outline{{background:#fff;border:1px solid {C['outline']};border-radius:6px;padding:7px 14px;font-size:12px;font-weight:600;cursor:pointer;color:#374151;text-decoration:none}}
.btn-outline:hover{{background:{C['surface_high']}}}
.btn-primary{{background:{C['dark']};color:#fff;border:none;border-radius:6px;padding:7px 14px;font-size:12px;font-weight:600;cursor:pointer;text-decoration:none;display:inline-block}}
.btn-primary:hover{{background:{C['mid']}}}
.alert{{padding:11px 16px;border-radius:8px;margin-bottom:16px;font-size:13px;font-weight:600}}
.alert.ok{{background:#e8f5e9;color:#1b5e20;border:1px solid #a5d6a7}}
.alert.err{{background:#ffdad6;color:{C['error']};border:1px solid #ffb4ab}}
.del-btn{{background:none;border:none;color:{C['error']};cursor:pointer;font-size:17px;padding:2px 5px;border-radius:4px}}
.del-btn:hover{{background:#ffdad6}}
.nav-details summary::-webkit-details-marker {{ display:none }}
.nav-details summary {{ list-style:none }}
.nav-details summary::marker {{ content:"" }}

/* ---- Chat Assistant (custom chat UI) ---- */
.chat-shell{{
    display:flex;
    flex-direction:column;
    height:calc(100vh - 110px);
    width:calc(100% - 48px);
    max-width:1400px;
    margin:24px auto;
    padding:0;
    background:#fff;
    border:1px solid {C['outline']};
    border-radius:16px;
    box-shadow:0 4px 24px rgba(0,0,0,.08);
    overflow:hidden;
}}
.chat-header{{
    display:flex;
    align-items:center;
    gap:10px;
    padding:16px 20px;
    border-bottom:1px solid {C['outline']};
    background:#fafbfc;
}}
.chat-header-avatar{{
    width:34px;height:34px;border-radius:50%;
    background:{C['dark']};color:#fff;
    display:flex;align-items:center;justify-content:center;
    font-size:14px;font-weight:700;flex-shrink:0;
}}
.chat-header-info{{display:flex;flex-direction:column;gap:1px}}
.chat-header-title{{font-size:14px;font-weight:700;color:{C['on_surface']}}}
.chat-header-status{{font-size:11px;color:#22c55e;display:flex;align-items:center;gap:5px}}
.chat-header-status::before{{content:"";width:6px;height:6px;border-radius:50%;background:#22c55e}}

#messages{{
    flex:1;
    overflow-y:auto;
    padding:20px;
    background:#f7f8fa;
    display:flex;
    flex-direction:column;
    gap:2px;
}}
#messages::-webkit-scrollbar{{width:6px}}
#messages::-webkit-scrollbar-thumb{{background:#d1d5db;border-radius:10px}}

.chat-empty-state{{
    margin:auto;
    text-align:center;
    color:#9ca3af;
    padding:40px 20px;
}}
.chat-empty-state-icon{{font-size:32px;margin-bottom:10px}}
.chat-empty-state-title{{font-size:14px;font-weight:600;color:#6b7280;margin-bottom:4px}}
.chat-empty-state-sub{{font-size:12px;color:#9ca3af}}

.msg-row{{display:flex;margin:5px 0;animation:msgIn .2s ease}}
@keyframes msgIn{{from{{opacity:0;transform:translateY(6px)}}to{{opacity:1;transform:translateY(0)}}}}
.msg-row.user-row{{justify-content:flex-end}}
.msg-row.bot-row{{justify-content:flex-start}}

.user,.bot{{
    display:inline-block;
    padding:10px 14px;
    border-radius:16px;
    width:-webkit-max-content;
    width:max-content;
    max-width:min(560px,75vw);
    font-size:13.5px;
    line-height:1.55;
    word-wrap:break-word;
    white-space:pre-wrap;
}}
.user{{
    background:{C['dark']};
    color:#fff;
    border-bottom-right-radius:4px;
}}
.bot{{
    background:#fff;
    color:#1f2937;
    border:1px solid #e5e7eb;
    border-bottom-left-radius:4px;
}}
.bot.typing{{color:#9ca3af;font-style:italic;white-space:nowrap}}
.bot.error{{background:#fff1f0;border-color:#ffccc7;color:{C['error']}}}

.msg-time{{
    font-size:10px;
    color:#b0b6bf;
    margin-top:3px;
    padding:0 4px;
}}
.msg-col{{display:inline-block;max-width:100%;vertical-align:top}}
.msg-col.user-col{{text-align:right}}
.msg-col.bot-col{{text-align:left}}
.msg-col .msg-time{{text-align:inherit}}

.chat-input{{
    display:flex;
    gap:10px;
    padding:14px 16px;
    border-top:1px solid {C['outline']};
    background:#fff;
}}
#msg{{
    flex:1;
    padding:11px 16px;
    border:1px solid #dfe3e8;
    border-radius:24px;
    font-size:13.5px;
    outline:none;
    transition:border-color .15s;
}}
#msg:focus{{border-color:{C['dark']}}}
.chat-input button{{
    padding:0 22px;
    background:{C['dark']};
    color:#fff;
    border:none;
    border-radius:24px;
    cursor:pointer;
    font-size:13px;
    font-weight:600;
    transition:background .15s,opacity .15s;
}}
.chat-input button:hover{{background:{C['mid']}}}
.chat-input button:disabled{{opacity:.5;cursor:not-allowed}}

.chat-empty{{
    padding:12px 16px;
    border:1px solid #ffe58f;
    border-radius:10px;
    background:#fffbe6;
    font-size:12.5px;
    color:#7a5c00;
    line-height:1.6;
    margin:14px 20px 0;
}}

/* ---- Chat history sidebar ---- */
.chat-layout{{
    display:flex;
    gap:18px;
    width:calc(100% - 48px);
    max-width:1420px;
    margin:20px auto;
    height:calc(100vh - 96px);
}}
.chat-layout .chat-shell{{margin:0;flex:1;min-width:0;height:100%}}
.chat-history-panel{{
    width:300px;
    flex-shrink:0;
    display:flex;
    flex-direction:column;
    background:#fff;
    border:1px solid {C['outline']};
    border-radius:16px;
    box-shadow:0 4px 24px rgba(0,0,0,.06);
    overflow:hidden;
}}
.chat-history-head{{
    padding:16px 16px 12px;
    border-bottom:1px solid {C['outline']};
    background:#fafbfc;
    display:flex;
    flex-direction:column;
    gap:10px;
}}
.chat-history-head-top{{
    display:flex;
    align-items:center;
    justify-content:space-between;
}}
.chat-history-head-title{{
    font-size:12px;
    font-weight:800;
    letter-spacing:.06em;
    text-transform:uppercase;
    color:{C['on_var']};
    display:flex;
    align-items:center;
    gap:6px;
}}
.chat-hist-count{{
    background:{C['surface_high']};
    color:{C['dark']};
    font-size:11px;
    font-weight:700;
    padding:2px 7px;
    border-radius:10px;
}}
.chat-clear-all-btn{{
    background:none;
    border:none;
    color:#9ca3af;
    font-size:11px;
    font-weight:600;
    cursor:pointer;
    padding:3px 6px;
    border-radius:6px;
    transition:all .15s ease;
    display:flex;
    align-items:center;
    gap:4px;
}}
.chat-clear-all-btn:hover{{
    color:{C['error']};
    background:#fee2e2;
}}
.chat-new-btn{{
    width:100%;
    padding:10px 14px;
    border:none;
    border-radius:10px;
    background:{C['dark']};
    color:#fff;
    font-size:12.5px;
    font-weight:700;
    cursor:pointer;
    display:flex;
    align-items:center;
    justify-content:center;
    gap:6px;
    transition:all .15s ease;
    box-shadow:0 2px 6px rgba(0,14,36,.1);
}}
.chat-new-btn:hover{{
    background:{C['mid']};
    transform:translateY(-1px);
    box-shadow:0 4px 10px rgba(0,14,36,.15);
}}
.chat-new-btn:active{{
    transform:translateY(0);
}}
.chat-search-wrap{{
    position:relative;
    width:100%;
}}
.chat-search-input{{
    width:100%;
    padding:7px 10px 7px 28px;
    font-size:12px;
    border:1px solid #e2e8f0;
    border-radius:8px;
    background:#fff;
    outline:none;
    transition:border-color .15s;
    color:{C['on_surface']};
}}
.chat-search-input:focus{{
    border-color:{C['dark']};
}}
.chat-search-ico{{
    position:absolute;
    left:8px;
    top:50%;
    transform:translateY(-50%);
    font-size:12px;
    color:#94a3b8;
    pointer-events:none;
}}
.chat-history-list{{
    flex:1;
    overflow-y:auto;
    padding:8px 8px;
    display:flex;
    flex-direction:column;
    gap:4px;
}}
.chat-history-list::-webkit-scrollbar{{width:5px}}
.chat-history-list::-webkit-scrollbar-thumb{{background:#e2e8f0;border-radius:6px}}

.chat-history-item{{
    display:flex;
    align-items:center;
    justify-content:space-between;
    width:100%;
    text-align:left;
    padding:9px 10px;
    border-radius:10px;
    background:transparent;
    cursor:pointer;
    border:1px solid transparent;
    transition:background .15s ease, border-color .15s ease, transform .2s ease, opacity .2s ease;
    position:relative;
    gap:8px;
    user-select:none;
}}
.chat-history-item:hover{{
    background:#f8fafc;
    border-color:#f1f5f9;
}}
.chat-history-item.active{{
    background:#f0f4ff;
    border-color:#c7d2fe;
}}
.chat-history-item.active .chat-history-item-title{{
    color:{C['dark']};
    font-weight:700;
}}
.chat-item-icon{{
    font-size:15px;
    flex-shrink:0;
    color:#64748b;
}}
.chat-history-item.active .chat-item-icon{{
    color:{C['dark']};
}}
.chat-item-content{{
    flex:1;
    min-width:0;
    display:flex;
    flex-direction:column;
    gap:2px;
}}
.chat-history-item-title{{
    font-size:12.5px;
    font-weight:600;
    color:{C['on_surface']};
    white-space:nowrap;
    overflow:hidden;
    text-overflow:ellipsis;
    line-height:1.3;
}}
.chat-history-item-meta{{
    display:flex;
    align-items:center;
    gap:6px;
    font-size:11px;
    color:#94a3b8;
}}
.chat-history-date{{
    white-space:nowrap;
    overflow:hidden;
    text-overflow:ellipsis;
}}
.chat-item-badge{{
    background:#e2e8f0;
    color:#475569;
    font-size:9.5px;
    font-weight:600;
    padding:1px 5px;
    border-radius:6px;
}}
.chat-item-del-btn{{
    width:26px;
    height:26px;
    border-radius:6px;
    border:none;
    background:transparent;
    color:#94a3b8;
    cursor:pointer;
    display:flex;
    align-items:center;
    justify-content:center;
    flex-shrink:0;
    opacity:0;
    transition:all .15s ease;
}}
.chat-history-item:hover .chat-item-del-btn,
.chat-history-item.active .chat-item-del-btn{{
    opacity:1;
}}
.chat-item-del-btn:hover{{
    background:#fee2e2;
    color:{C['error']};
    transform:scale(1.1);
}}
.chat-item-del-btn:active{{
    transform:scale(0.95);
}}
.chat-history-item.deleting{{
    opacity:0;
    transform:translateX(-14px);
    max-height:0;
    margin:0;
    padding:0;
    border:none;
    pointer-events:none;
    overflow:hidden;
}}
.chat-history-empty{{
    padding:30px 14px;
    text-align:center;
    color:#94a3b8;
    font-size:12px;
    display:flex;
    flex-direction:column;
    align-items:center;
    gap:6px;
}}
.chat-history-empty-icon{{
    font-size:24px;
    opacity:.6;
}}
@media(max-width:900px){{
  .chat-layout{{flex-direction:column;height:auto}}
  .chat-history-panel{{width:100%;max-height:260px}}
}}
@media(max-width:800px){{
  .chat-shell{{height:calc(100vh - 90px);margin:12px auto;border-radius:12px}}
}}
@media(max-width:800px){{
  .sidebar{{width:200px;top:0;bottom:0;height:100vh;overflow-y:auto;transform:translateX(-100%);transition:transform .25s ease;box-shadow:2px 0 14px rgba(0,0,0,.18)}}
  .sidebar.open{{transform:translateX(0)}}
  .main{{margin-left:0}}
  .topbar{{left:0;padding:0 12px}}
  .topbar-title{{display:none}}
  .topbar-nav{{display:none}}
  .search-box{{width:120px}}
  .hamburger{{display:block}}
  .nav-overlay.show{{display:block}}
  .chat-shell{{height:calc(100vh - 90px);margin:12px auto;padding:0 10px}}
}}
</style>"""

def _html_head(title="Monarch Surveyors") -> str:
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<link rel="icon" href="{LOGO_URL}" type="image/jpeg">
<title>{title} | Monarch Surveyors</title>
{SHARED_CSS}
<script>
/* ── Cross-tab filter persistence ──
   Filters set on a page survive navigating to another tab and back.
   A hard refresh (F5 / reload) always wipes them so the page opens clean. */
window.MonarchFilters = {{
  _isReload: function() {{
    try {{
      var nav = performance.getEntriesByType('navigation')[0];
      if (nav) return nav.type === 'reload';
      if (performance.navigation) return performance.navigation.type === 1;
    }} catch (e) {{}}
    return false;
  }},
  _key: function(pageKey) {{ return 'monarch_filter_' + pageKey; }},
  clear: function(pageKey) {{ sessionStorage.removeItem(this._key(pageKey)); }},
  save: function(pageKey, obj) {{ sessionStorage.setItem(this._key(pageKey), JSON.stringify(obj)); }},
  load: function(pageKey) {{
    try {{
      var raw = sessionStorage.getItem(this._key(pageKey));
      return raw ? JSON.parse(raw) : null;
    }} catch (e) {{ return null; }}
  }},
  /* Call once per page on load. Returns null (and wipes storage) on a hard
     refresh; otherwise returns the saved filter object (or null if none). */
  init: function(pageKey) {{
    if (this._isReload()) {{ this.clear(pageKey); return null; }}
    return this.load(pageKey);
  }}
}};
</script>
<script>
/* ── Real-time data sync ──
   Polls /api/data_version. The counter bumps the instant ANY user saves,
   edits, uploads, or deletes anything anywhere in the app. When it changes,
   this tab refreshes itself automatically — no manual reload needed.
   Guarded so it never yanks the page out from under someone mid-edit: if a
   text field/textarea/select is focused, or a modal/dialog looks open, the
   refresh is deferred until they're done. */
(function() {{
  var baseline = null;
  var pending  = false;

  function isEditing() {{
    var el = document.activeElement;
    if (el && /^(INPUT|TEXTAREA|SELECT)$/.test(el.tagName)) {{
      var type = (el.getAttribute('type') || '').toLowerCase();
      if (el.tagName !== 'INPUT' || !['button','submit','checkbox','radio'].includes(type)) return true;
    }}
    if (el && el.isContentEditable) return true;
    if (window._dirty || document.getElementById('main-project-form')) return true;
    if (document.querySelector('.modal.open, .modal.show, [role="dialog"][open], dialog[open]')) return true;
    return false;
  }}

  function poll() {{
    fetch('/api/data_version', {{ cache: 'no-store' }})
      .then(function(r) {{ return r.ok ? r.json() : null; }})
      .then(function(d) {{
        if (!d) return;
        if (baseline === null) {{ baseline = d.v; return; }}
        if (d.v !== baseline) {{
          if (isEditing()) {{ pending = true; return; }}
          window.location.reload();
        }}
      }})
      .catch(function() {{}});
  }}

  document.addEventListener('focusout', function() {{
    if (pending && !isEditing()) {{ window.location.reload(); }}
  }});

  poll();
  setInterval(poll, 4000);
}})();
</script>"""

def _sidebar(active="dashboard", sub_section="", is_admin=None) -> str:
    if is_admin is None:
        is_admin = getattr(_req_ctx, "is_admin", False)
    role = getattr(_req_ctx, "role", "")
    if role == "director":
        nav_items = [i for i in NAV_ITEMS if i[0] in DIRECTOR_NAV_KEYS]
    elif role == "accounts":
        nav_items = [i for i in NAV_ITEMS if i[0] in ACCOUNTS_NAV_KEYS]
    elif role == "procurement":
        nav_items = [i for i in NAV_ITEMS if i[0] in PROCUREMENT_NAV_KEYS]
    elif role in ("pm", "head"):
        nav_items = [i for i in NAV_ITEMS if i[0] not in PM_HIDDEN_NAV_KEYS]
    else:
        nav_items = NAV_ITEMS
    links = ""
    for key, icon, label in nav_items:
        cls = "nav-item active" if key == active else "nav-item"
        links += f'<a href="/{key}" class="{cls}"><span class="nav-icon">{icon}</span>{label}</a>\n'

    if is_admin:
        cls = "nav-item active" if active == "user_management" else "nav-item"
        links += (f'<a href="/user_management" class="{cls}">'
                  f'<span class="nav-icon">👤</span>User Management</a>\n')

    return f"""
<aside class="sidebar">
  <div class="brand" style="display:flex;align-items:center;gap:10px;padding-bottom:14px">
    <img src="{LOGO_URL}" alt="Monarch logo" style="width:36px;height:36px;border-radius:6px;flex-shrink:0">
    <span style="line-height:1.3">Monarch<br>Surveyors &amp; Engineering<br>Consultants Ltd.</span>
  </div>
  {links}
  <button class="sidebar-btn" onclick="window.location='/reports'">⊕ New Report</button>
</aside>
<div class="nav-overlay" onclick="document.querySelector('.sidebar').classList.remove('open');this.classList.remove('show')"></div>"""


def _payments_sidebar(active="payments_dashboard") -> str:
    """Sidebar for the standalone Payments portal (Weekly Payment Booking +
    Payment Approvals). Kept visually separate from the main system sidebar —
    reached via the portal switch on the login page."""
    role = getattr(_req_ctx, "role", "")
    nav_items = PAYMENTS_NAV_ITEMS
    # PM/Head only ever create bookings — they don't review/approve, so hide
    # the approvals link for them (matches their existing permissions).
    if role in ("pm", "head"):
        nav_items = [i for i in PAYMENTS_NAV_ITEMS if i[0] != "payment_approvals"]

    links = ""
    for key, icon, label in nav_items:
        cls = "nav-item active" if key == active else "nav-item"
        links += f'<a href="/{key}" class="{cls}"><span class="nav-icon">{icon}</span>{label}</a>\n'

    return f"""
<aside class="sidebar">
  <div class="brand" style="display:flex;align-items:center;gap:10px;padding-bottom:14px">
    <img src="{LOGO_URL}" alt="Monarch logo" style="width:36px;height:36px;border-radius:6px;flex-shrink:0">
    <span style="line-height:1.3">Monarch Payments<br><span style="font-weight:400;font-size:11px;color:#9aa0a6">Booking &amp; Approvals Portal</span></span>
  </div>
  {links}
</aside>
<div class="nav-overlay" onclick="document.querySelector('.sidebar').classList.remove('open');this.classList.remove('show')"></div>"""


def _topbar(subtitle="FINANCIAL MONITORING SYSTEM", username="Admin", search_value="", portal="main") -> str:
    initials = "".join(w[0] for w in (username or "A").split()[:2]).upper() or "A"
    sv = (search_value or "").replace('"', "&quot;")
    return f"""
<header class="topbar">
  <button class="hamburger" aria-label="Menu" onclick="document.querySelector('.sidebar').classList.toggle('open');document.querySelector('.nav-overlay').classList.toggle('show')">☰</button>
  <span class="topbar-title">{subtitle}</span>
  <div class="topbar-right">
    <form method="get" action="/search" style="margin:0">
      <input class="search-box" name="q" placeholder="🔍  Search…" type="text" value="{sv}" autocomplete="off">
    </form>
    <nav class="topbar-nav">
      <a href="/dashboard">Dashboard</a>
      <a href="/reports">Reports</a>
      <a href="/kpi">KPI</a>
    </nav>
    <div class="tb-portal-switch">
      <a href="/dashboard" class="tb-portal-btn{' active' if portal=='main' else ''}">🏢 Project Master</a>
      <a href="/payments_dashboard" class="tb-portal-btn{' active' if portal=='payments' else ''}">💳 Payment Booking</a>
    </div>
    <button class="proj-sel-btn" onclick="window.location='/data_entry'">Project Selector</button>
    <span style="font-size:18px;cursor:pointer">🔔</span>
    <a class="avatar-link" href="/logout" title="Sign out">
      <div class="avatar">{initials}</div>
      <span style="font-size:13px;font-weight:600">{username}</span>
    </a>
  </div>
</header>"""

def fetch_chat_sessions(username: str = "") -> list[dict]:
    """All saved chat sessions, most recently updated first. Supports both
    public.chat_sessions and public.n8n_chat_histories automatically.
    """
    cols = ["id", "username", "title", "created_at", "updated_at", "msg_count"]
    try:
        with ENGINE.connect() as conn:
            # 1. Try public.chat_sessions first
            try:
                rows = []
                if username:
                    rows = conn.execute(text("""
                        SELECT cs.id, cs.username, cs.title, cs.created_at, cs.updated_at,
                               COUNT(cm.id) as msg_count
                        FROM public.chat_sessions cs
                        LEFT JOIN public.chat_messages cm ON cm.session_id = cs.id
                        WHERE LOWER(cs.username) = LOWER(:u)
                        GROUP BY cs.id, cs.username, cs.title, cs.created_at, cs.updated_at
                        ORDER BY COALESCE(cs.updated_at, cs.created_at) DESC
                    """), {"u": username}).fetchall()
                if not rows:
                    rows = conn.execute(text("""
                        SELECT cs.id, cs.username, cs.title, cs.created_at, cs.updated_at,
                               COUNT(cm.id) as msg_count
                        FROM public.chat_sessions cs
                        LEFT JOIN public.chat_messages cm ON cm.session_id = cs.id
                        GROUP BY cs.id, cs.username, cs.title, cs.created_at, cs.updated_at
                        ORDER BY COALESCE(cs.updated_at, cs.created_at) DESC
                    """)).fetchall()
                if rows:
                    return [dict(zip(cols, r)) for r in rows]
            except Exception:
                pass

            # 2. Fallback to public.n8n_chat_histories
            try:
                import json as _json
                rows = conn.execute(text(
                    "SELECT session_id, message, id FROM public.n8n_chat_histories ORDER BY id ASC"
                )).fetchall()
                sess_map = {}  # sid -> dict
                for sid, msg, mid in rows:
                    sid = str(sid or '').strip().rstrip('\n').rstrip('\r').strip()
                    if not sid:
                        continue
                    if isinstance(msg, str):
                        try: msg_obj = _json.loads(msg)
                        except: msg_obj = {}
                    else:
                        msg_obj = msg or {}
                    mtype = str(msg_obj.get('type', '')).lower()
                    content = str(msg_obj.get('content', '') or '').strip()
                    if mtype == 'human':
                        if content.startswith('User Question:'):
                            content = content.split('User Question:', 1)[1].split('\n\nDatabase Results:', 1)[0].strip()
                        elif "The user's request is:" in content:
                            parts = content.split("The user's request is:", 1)
                            if len(parts) > 1:
                                content = parts[1].split("Rules:", 1)[0].split("\n\n", 1)[0].strip()
                        is_sql = any(kw in content.upper() for kw in ['SELECT ', 'INSERT ', 'UPDATE ', 'DELETE ', 'WITH ', '```SQL'])
                        if content and not is_sql:
                            clean_c = " ".join(content.split())
                            title = (clean_c[:48] + '...') if len(clean_c) > 48 else clean_c
                            if sid not in sess_map:
                                sess_map[sid] = {
                                    "id": sid,
                                    "username": username or "User",
                                    "title": title,
                                    "created_at": None,
                                    "updated_at": None,
                                    "msg_count": 0
                                }
                    if sid in sess_map:
                        sess_map[sid]["msg_count"] = sess_map[sid].get("msg_count", 0) + 1
                    elif sid not in sess_map:
                        sess_map[sid] = {
                            "id": sid,
                            "username": username or "User",
                            "title": f"Chat {sid[:8]}",
                            "created_at": None,
                            "updated_at": None,
                            "msg_count": 1
                        }

                return list(reversed(list(sess_map.values())))
            except Exception as e:
                print(f"[fetch_chat_sessions n8n fallback] {e}")
        return []
    except Exception as e:
        print(f"[fetch_chat_sessions] {e}")
        return []

def delete_chat_session(session_id: str, username: str = "") -> bool:
    """Delete a chat session and all associated messages."""
    if not session_id:
        return False
    try:
        sid_clean = str(session_id).strip().rstrip('\n').rstrip('\r').strip()
        with ENGINE.begin() as conn:
            # 1. Delete from chat_messages and chat_sessions
            try:
                conn.execute(text(
                    "DELETE FROM public.chat_messages WHERE session_id = :sid"
                ), {"sid": sid_clean})
                conn.execute(text(
                    "DELETE FROM public.chat_sessions WHERE id = :sid"
                ), {"sid": sid_clean})
            except Exception as e:
                print(f"[delete_chat_session chat_sessions] {e}")
            
            # 2. Delete from n8n_chat_histories
            try:
                conn.execute(text(
                    "DELETE FROM public.n8n_chat_histories "
                    "WHERE regexp_replace(session_id, '[\\s]+', '', 'g') = regexp_replace(:sid, '[\\s]+', '', 'g')"
                ), {"sid": sid_clean})
            except Exception as e:
                print(f"[delete_chat_session n8n_chat_histories] {e}")
        return True
    except Exception as e:
        print(f"[delete_chat_session] {e}")
        return False

def clear_all_chat_history(username: str = "") -> bool:
    """Clear all chat sessions and messages."""
    try:
        with ENGINE.begin() as conn:
            try:
                if username:
                    conn.execute(text("""
                        DELETE FROM public.chat_messages WHERE session_id IN (
                            SELECT id FROM public.chat_sessions WHERE LOWER(username)=LOWER(:u)
                        )
                    """), {"u": username})
                    conn.execute(text(
                        "DELETE FROM public.chat_sessions WHERE LOWER(username)=LOWER(:u)"
                    ), {"u": username})
                else:
                    conn.execute(text("DELETE FROM public.chat_messages"))
                    conn.execute(text("DELETE FROM public.chat_sessions"))
            except Exception:
                pass
            try:
                conn.execute(text("DELETE FROM public.n8n_chat_histories"))
            except Exception:
                pass
        return True
    except Exception as e:
        print(f"[clear_all_chat_history] {e}")
        return False

def fetch_chat_messages(session_id) -> list[dict]:
    cols = ["id", "session_id", "who", "message", "created_at"]
    try:
        with ENGINE.connect() as conn:
            # 1. Try public.chat_messages first
            try:
                rows = conn.execute(text(
                    "SELECT id, session_id, who, message, created_at FROM public.chat_messages "
                    "WHERE session_id=:sid ORDER BY id ASC"
                ), {"sid": session_id}).fetchall()
                if rows:
                    return [dict(zip(cols, r)) for r in rows]
            except Exception:
                pass

            # 2. Fallback to public.n8n_chat_histories
            try:
                import json as _json
                sid_clean = str(session_id).strip().rstrip('\n').rstrip('\r').strip()
                rows = conn.execute(text(
                    "SELECT id, session_id, message FROM public.n8n_chat_histories "
                    "WHERE regexp_replace(session_id, '[\\s]+', '', 'g')=regexp_replace(:sid, '[\\s]+', '', 'g') ORDER BY id ASC"
                ), {"sid": sid_clean}).fetchall()
                msgs = []
                for mid, sid, msg in rows:
                    if isinstance(msg, str):
                        try: msg_obj = _json.loads(msg)
                        except: msg_obj = {}
                    else:
                        msg_obj = msg or {}
                    # n8n stores messages with top-level 'type' and 'content' fields
                    mtype = str(msg_obj.get('type', '')).lower()
                    content = str(msg_obj.get('content', '') or '').strip()
                    if mtype == 'human':
                        who = 'user'
                        # Format 1 (Answer Maker): "User Question:\n<question>\n\nDatabase Results:"
                        if content.startswith('User Question:'):
                            content = content.split('User Question:', 1)[1].split('\n\nDatabase Results:', 1)[0].strip()
                        # Format 2 (SQL Generator): "You are an expert...The user's request is:\n<question>\n"
                        elif "The user's request is:" in content:
                            parts = content.split("The user's request is:", 1)
                            if len(parts) > 1:
                                content = parts[1].split("Rules:", 1)[0].split("\n\n", 1)[0].strip()
                        # Skip if content is still SQL
                        is_sql = any(kw in content.upper() for kw in ['SELECT ', 'INSERT ', 'UPDATE ', 'DELETE ', 'WITH ', '```SQL'])
                        if is_sql or not content:
                            continue
                    elif mtype == 'ai':
                        who = 'bot'
                        # Skip raw SQL bot responses (old sessions)
                        is_sql = any(kw in content.upper() for kw in ['SELECT ', 'INSERT ', 'UPDATE ', 'DELETE ', 'WITH ', '```SQL'])
                        if is_sql or not content:
                            continue
                    else:
                        continue  # skip tool/system messages
                    msgs.append({
                        "id": mid,
                        "session_id": sid,
                        "who": who,
                        "message": content,
                        "created_at": None
                    })
                return msgs
            except Exception as e:
                print(f"[fetch_chat_messages n8n fallback] {e}")
        return []
    except Exception as e:
        print(f"[fetch_chat_messages] {e}")
        return []

def chat_page(username: str = "Admin") -> str:
    request_path = getattr(_req_ctx, "path", "") or ""
    chat_url = (parse_qs(urlparse(request_path).query).get("url", [N8N_CHAT_URL])[0] or N8N_CHAT_URL).strip()
    safe_chat_url = html.escape(chat_url, quote=True)

    warning_html = ""
    if not chat_url:
        warning_html = (
            '<div class="chat-empty">No n8n chat URL is configured yet. '
            'Set the <strong>N8N_CHAT_URL</strong> environment variable or open this page with '
            '<code>?url=https://your-n8n-url</code> to load your bot.</div>'
        )

    def _fmt_hist_date(v):
        if not v:
            return ""
        s = str(v).replace("T", " ")[:16]
        return s

    sessions = fetch_chat_sessions(username)
    sess_count = len(sessions)
    if sessions:
        sess_items = ""
        for s in sessions:
            sid   = s.get("id")
            title = html.escape((s.get("title") or f"Chat {sid}"), quote=True)
            raw_when = s.get("updated_at") or s.get("created_at")
            when_fmt = _fmt_hist_date(raw_when)
            when_str = html.escape(when_fmt if when_fmt else "Saved chat", quote=True)
            msg_count = s.get("msg_count", 0)
            badge_html = f'<span class="chat-item-badge">{msg_count} msgs</span>' if msg_count > 1 else ''
            sess_items += (
                f'<div class="chat-history-item" data-session-id="{sid}" onclick="loadSession(\'{sid}\', this)">'
                f'  <div class="chat-item-icon">💬</div>'
                f'  <div class="chat-item-content">'
                f'    <div class="chat-history-item-title" title="{title}">{title}</div>'
                f'    <div class="chat-history-item-meta">'
                f'      <span class="chat-history-date">{when_str}</span>'
                f'      {badge_html}'
                f'    </div>'
                f'  </div>'
                f'  <button type="button" class="chat-item-del-btn" title="Delete conversation" '
                f'onclick="deleteSession(event, \'{sid}\', this)">'
                f'    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
                f'      <polyline points="3 6 5 6 21 6"></polyline>'
                f'      <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>'
                f'      <line x1="10" y1="11" x2="10" y2="17"></line>'
                f'      <line x1="14" y1="11" x2="14" y2="17"></line>'
                f'    </svg>'
                f'  </button>'
                f'</div>'
            )
    else:
        sess_items = (
            '<div class="chat-history-empty" id="historyEmptyState">'
            '<div class="chat-history-empty-icon">💬</div>'
            '<div>No saved conversations yet.</div>'
            '<div style="font-size:11px;color:#94a3b8">Start a chat to see history</div>'
            '</div>'
        )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Chat Assistant | Monarch Surveyors</title>
{SHARED_CSS}
</head>
<body>
{_sidebar("chat", is_admin=getattr(_req_ctx, "is_admin", False))}
{_topbar("CHAT ASSISTANT", username, search_value="")}

<div class="main">

<div class="chat-layout">

  <div class="chat-history-panel">
    <div class="chat-history-head">
      <div class="chat-history-head-top">
        <div class="chat-history-head-title">
          <span>Chat History</span>
          <span class="chat-hist-count" id="histCount">{sess_count}</span>
        </div>
        <button type="button" class="chat-clear-all-btn" id="clearAllBtn" onclick="confirmClearAll()" title="Clear all chat history" {'style="display:none"' if not sess_count else ''}>
          🗑️ Clear
        </button>
      </div>
      <button type="button" class="chat-new-btn" onclick="newChat()">＋ New Chat</button>
      <div class="chat-search-wrap">
        <span class="chat-search-ico">🔍</span>
        <input type="text" id="chatSearch" class="chat-search-input" placeholder="Search chats..." oninput="filterHistory(this.value)">
      </div>
    </div>
    <div class="chat-history-list" id="historyList">
      {sess_items}
      <div class="chat-history-empty" id="historyEmptySearch" style="display:none">
        <div class="chat-history-empty-icon">🔍</div>
        <div>No matching chats found</div>
      </div>
    </div>
  </div>

  <div class="chat-shell">

    <div class="chat-header">
        <div class="chat-header-avatar">M</div>
        <div class="chat-header-info">
            <div class="chat-header-title">Monarch Assistant</div>
            <div class="chat-header-status">Online</div>
        </div>
    </div>

    {warning_html}

    <div id="messages">
        <div class="chat-empty-state" id="emptyState">
            <div class="chat-empty-state-icon">💬</div>
            <div class="chat-empty-state-title">Ask me anything about your projects</div>
            <div class="chat-empty-state-sub">e.g. "How many active projects does Sopan manage?"</div>
        </div>
    </div>

    <div class="chat-input">
        <input
            id="msg"
            type="text"
            placeholder="Ask anything..."
            onkeydown="if(event.key==='Enter') sendMessage()"
            autocomplete="off"
        >
        <button id="sendBtn" onclick="sendMessage()">Send</button>
    </div>

  </div>

</div>

</div>

<script>

const CHAT_URL = "{safe_chat_url}";
let sessionId = "web-" + Date.now() + "-" + Math.random().toString(36).slice(2);

function timeNow(){{
    const d = new Date();
    return d.toLocaleTimeString([], {{ hour: '2-digit', minute: '2-digit' }});
}}

function escapeHtml(str){{
    const div = document.createElement("div");
    div.textContent = str || "";
    return div.innerHTML;
}}

function addMessage(text, who, whenStr){{
    const emptyState = document.getElementById("emptyState");
    if(emptyState) emptyState.remove();

    const messagesEl = document.getElementById("messages");
    const row = document.createElement("div");
    row.className = "msg-row " + (who === "user" ? "user-row" : "bot-row");

    const col = document.createElement("div");
    col.className = "msg-col " + (who === "user" ? "user-col" : "bot-col");

    const bubble = document.createElement("div");
    bubble.className = who;
    bubble.innerHTML = escapeHtml(text);

    const time = document.createElement("div");
    time.className = "msg-time";
    time.textContent = whenStr || timeNow();

    col.appendChild(bubble);
    col.appendChild(time);
    row.appendChild(col);
    messagesEl.appendChild(row);
    messagesEl.scrollTop = messagesEl.scrollHeight;

    return bubble;
}}

function fmtHistTime(iso){{
    if(!iso) return "";
    const d = new Date(iso);
    if(isNaN(d.getTime())) return "";
    return d.toLocaleTimeString([], {{ hour: '2-digit', minute: '2-digit' }});
}}

function clearMessages(){{
    const messagesEl = document.getElementById("messages");
    messagesEl.innerHTML =
        '<div class="chat-empty-state" id="emptyState">' +
        '<div class="chat-empty-state-icon">💬</div>' +
        '<div class="chat-empty-state-title">Ask me anything about your projects</div>' +
        '<div class="chat-empty-state-sub">e.g. "How many active projects does Sopan manage?"</div>' +
        '</div>';
}}

async function loadSession(sid, itemEl){{
    document.querySelectorAll('.chat-history-item').forEach(function(b){{ b.classList.remove('active'); }});
    if(itemEl) itemEl.classList.add('active');

    clearMessages();
    const messagesEl = document.getElementById("messages");
    messagesEl.innerHTML = '<div class="chat-empty-state"><div class="chat-empty-state-title">Loading conversation…</div></div>';

    try {{
        const resp = await fetch('/chat_history?session_id=' + encodeURIComponent(sid));
        const data = await resp.json();
        const msgs = data.messages || [];
        clearMessages();
        if (!msgs.length) {{
            document.getElementById("messages").innerHTML =
                '<div class="chat-empty-state"><div class="chat-empty-state-icon">💬</div>' +
                '<div class="chat-empty-state-title">No messages in this conversation yet</div></div>';
        }} else {{
            document.getElementById("emptyState") && document.getElementById("emptyState").remove();
            msgs.forEach(function(m){{
                const who = (m.who || '').toLowerCase() === 'user' ? 'user' : 'bot';
                addMessage(m.message || '', who, fmtHistTime(m.created_at));
            }});
        }}
        sessionId = String(sid);
    }} catch(err) {{
        clearMessages();
        addMessage("Couldn't load this conversation.", "bot");
        console.error("Load history error:", err);
    }}
}}

function newChat(){{
    document.querySelectorAll('.chat-history-item').forEach(function(b){{ b.classList.remove('active'); }});
    sessionId = "web-" + Date.now() + "-" + Math.random().toString(36).slice(2);
    clearMessages();
    const input = document.getElementById("msg");
    if(input) input.focus();
}}

async function deleteSession(event, sid, btnEl){{
    if (event) {{
        event.stopPropagation();
        event.preventDefault();
    }}
    if (!sid) return;
    if (!confirm("Are you sure you want to delete this chat conversation?")) return;

    const itemEl = btnEl ? btnEl.closest('.chat-history-item') : document.querySelector(`.chat-history-item[data-session-id="${{sid}}"]`);
    if (itemEl) {{
        itemEl.classList.add('deleting');
    }}

    try {{
        const resp = await fetch('/delete_chat_session', {{
            method: 'POST',
            headers: {{ 'Content-Type': 'application/json' }},
            body: JSON.stringify({{ session_id: sid }})
        }});
        const data = await resp.json();
        if (data.ok) {{
            setTimeout(function(){{
                if (itemEl) itemEl.remove();
                if (sessionId === sid) {{
                    newChat();
                }}
                updateHistoryCount();
            }}, 220);
        }} else {{
            if (itemEl) itemEl.classList.remove('deleting');
            alert(data.error || "Could not delete conversation.");
        }}
    }} catch (err) {{
        console.error("Delete error:", err);
        if (itemEl) itemEl.classList.remove('deleting');
        alert("Failed to delete chat session.");
    }}
}}

async function confirmClearAll(){{
    const items = document.querySelectorAll('.chat-history-item');
    if (!items.length) return;
    if (!confirm("Clear all conversations from history? This cannot be undone.")) return;

    try {{
        const resp = await fetch('/clear_chat_history', {{
            method: 'POST',
            headers: {{ 'Content-Type': 'application/json' }}
        }});
        const data = await resp.json();
        if (data.ok) {{
            const listEl = document.getElementById('historyList');
            listEl.innerHTML =
                '<div class="chat-history-empty" id="historyEmptyState">' +
                '<div class="chat-history-empty-icon">💬</div>' +
                '<div>No saved conversations yet.</div>' +
                '<div style="font-size:11px;color:#94a3b8">Start a chat to see history</div>' +
                '</div>' +
                '<div class="chat-history-empty" id="historyEmptySearch" style="display:none">' +
                '<div class="chat-history-empty-icon">🔍</div>' +
                '<div>No matching chats found</div>' +
                '</div>';
            newChat();
            updateHistoryCount();
        }} else {{
            alert(data.error || "Could not clear chat history.");
        }}
    }} catch (err) {{
        console.error("Clear all error:", err);
        alert("Failed to clear chat history.");
    }}
}}

function updateHistoryCount(){{
    const items = document.querySelectorAll('.chat-history-item');
    const countEl = document.getElementById('histCount');
    const clearBtn = document.getElementById('clearAllBtn');
    const emptyState = document.getElementById('historyEmptyState');

    if (countEl) countEl.textContent = items.length;
    if (clearBtn) clearBtn.style.display = items.length > 0 ? 'flex' : 'none';

    if (items.length === 0 && !emptyState) {{
        const listEl = document.getElementById('historyList');
        const emptyDiv = document.createElement('div');
        emptyDiv.className = 'chat-history-empty';
        emptyDiv.id = 'historyEmptyState';
        emptyDiv.innerHTML = '<div class="chat-history-empty-icon">💬</div><div>No saved conversations yet.</div><div style="font-size:11px;color:#94a3b8">Start a chat to see history</div>';
        listEl.prepend(emptyDiv);
    }}
}}

function filterHistory(query){{
    const q = (query || '').toLowerCase().trim();
    const items = document.querySelectorAll('.chat-history-item');
    let visibleCount = 0;
    items.forEach(function(item){{
        const title = (item.querySelector('.chat-history-item-title')?.textContent || '').toLowerCase();
        const date = (item.querySelector('.chat-history-date')?.textContent || '').toLowerCase();
        if (!q || title.includes(q) || date.includes(q)) {{
            item.style.display = 'flex';
            visibleCount++;
        }} else {{
            item.style.display = 'none';
        }}
    }});
    const emptySearch = document.getElementById('historyEmptySearch');
    if (emptySearch) {{
        emptySearch.style.display = (visibleCount === 0 && items.length > 0 && q.length > 0) ? 'flex' : 'none';
    }}
}}

async function refreshHistoryList(){{
    try {{
        const resp = await fetch('/chat_sessions');
        const data = await resp.json();
        const sessions = data.sessions || [];
        const listEl = document.getElementById('historyList');
        if (!listEl) return;

        if (!sessions.length) {{
            listEl.innerHTML =
                '<div class="chat-history-empty" id="historyEmptyState">' +
                '<div class="chat-history-empty-icon">💬</div>' +
                '<div>No saved conversations yet.</div>' +
                '<div style="font-size:11px;color:#94a3b8">Start a chat to see history</div>' +
                '</div>';
            updateHistoryCount();
            return;
        }}

        let html = '';
        sessions.forEach(function(s){{
            const sid = escapeHtml(s.id || '');
            const title = escapeHtml(s.title || ('Chat ' + sid));
            const whenStr = escapeHtml(s.updated_at || s.created_at || 'Saved chat');
            const isActive = (sessionId === s.id) ? ' active' : '';
            const msgCount = s.msg_count || 0;
            const badge = msgCount > 1 ? `<span class="chat-item-badge">${{msgCount}} msgs</span>` : '';

            html +=
                `<div class="chat-history-item${{isActive}}" data-session-id="${{sid}}" onclick="loadSession('${{sid}}', this)">` +
                `  <div class="chat-item-icon">💬</div>` +
                `  <div class="chat-item-content">` +
                `    <div class="chat-history-item-title" title="${{title}}">${{title}}</div>` +
                `    <div class="chat-history-item-meta">` +
                `      <span class="chat-history-date">${{whenStr}}</span>` +
                `      ${{badge}}` +
                `    </div>` +
                `  </div>` +
                `  <button type="button" class="chat-item-del-btn" title="Delete conversation" onclick="deleteSession(event, '${{sid}}', this)">` +
                `    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">` +
                `      <polyline points="3 6 5 6 21 6"></polyline>` +
                `      <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>` +
                `      <line x1="10" y1="11" x2="10" y2="17"></line>` +
                `      <line x1="14" y1="11" x2="14" y2="17"></line>` +
                `    </svg>` +
                `  </button>` +
                `</div>`;
        }});

        html += '<div class="chat-history-empty" id="historyEmptySearch" style="display:none"><div class="chat-history-empty-icon">🔍</div><div>No matching chats found</div></div>';
        listEl.innerHTML = html;
        updateHistoryCount();
        const searchInput = document.getElementById('chatSearch');
        if (searchInput && searchInput.value) filterHistory(searchInput.value);
    }} catch (e) {{
        console.error("Refresh history error:", e);
    }}
}}

async function sendMessage(){{
    const input = document.getElementById("msg");
    const sendBtn = document.getElementById("sendBtn");
    const text = input.value.trim();

    if(!text) return;
    if(!CHAT_URL){{
        addMessage("Chat URL not configured.", "bot");
        return;
    }}

    addMessage(text, "user");
    input.value = "";
    input.disabled = true;
    sendBtn.disabled = true;

    const typingBubble = addMessage("Typing...", "bot");
    typingBubble.classList.add("typing");

    try {{
        const response = await fetch(CHAT_URL, {{
            method: "POST",
            headers: {{ "Content-Type": "application/json" }},
            body: JSON.stringify({{ chatInput: text, sessionId: sessionId }})
        }});

        if(!response.ok) throw new Error("HTTP " + response.status);

        let answer = "No response received.";
        try {{
            const textResp = await response.text();
            try {{
                const data = JSON.parse(textResp);
                answer = data.output || data.text || data.response || data.answer || (typeof data === "string" ? data : JSON.stringify(data));
            }} catch (jsonErr) {{
                if (textResp && textResp.trim()) answer = textResp.trim();
            }}
        }} catch (readErr) {{
            console.error("Response error:", readErr);
        }}

        typingBubble.classList.remove("typing");
        typingBubble.innerHTML = escapeHtml(answer);

        // Refresh history list so the new chat session appears in sidebar
        setTimeout(refreshHistoryList, 500);

    }} catch (err) {{
        typingBubble.classList.remove("typing");
        typingBubble.classList.add("error");
        typingBubble.innerHTML = "Couldn't reach the assistant. Please try again.";
        console.error("Chat error:", err);
    }} finally {{
        input.disabled = false;
        sendBtn.disabled = false;
        input.focus();
    }}
}}

</script>
</body>
</html>"""

def _sel(name: str, options: list[str], selected: str = "", placeholder: str = "Select…",
         cls: str = "", onchange: str = "", extra: str = "") -> str:
    opts = f'<option value="" disabled {"selected" if not selected else ""}>{placeholder}</option>'
    for o in options:
        opts += f'<option value="{o}" {"selected" if o == selected else ""}>{o}</option>'
    return (f'<select name="{name}" class="{cls}"'
            + (f' onchange="{onchange}"' if onchange else "")
            + (f' {extra}' if extra else "") + f'>{opts}</select>')

def _field_html(label: str, input_html: str) -> str:
    return f'<div class="field"><label>{label}</label>{input_html}</div>'

def _fmt_inr(n) -> str:
    try:    return f"{int(abs(float(n))):,}"
    except: return "0"


# ═══════════════════════════════════════════════════════════════
#  LOGIN PAGE
# ═══════════════════════════════════════════════════════════════

def login_page(alert_html: str = "", next_path: str = "/dashboard") -> str:
    year = datetime.now().year
    try:
        projects    = fetch_all_projects()
        lp_count    = len(projects)
        lp_ov       = sum(_rev_ov(p)                         for p in projects)
        lp_bil      = sum(_n(p.get("billed_amount"))         for p in projects)
        lp_col      = sum(_n(p.get("collection_till_date"))  for p in projects)
        lp_ub       = max(lp_bil - lp_col, 0)
        lp_pct      = f"{lp_col / lp_ov * 100:.1f}%" if lp_ov else "—"
        lp_ov_fmt   = f"₹ {lp_ov  / 1e7:.2f} Crore" if lp_ov  else "—"
        lp_col_fmt  = f"{lp_pct} (₹ {lp_col / 1e7:.1f} Cr)" if lp_col else "—"
        lp_ub_fmt   = f"₹ {lp_ub  / 1e7:.2f} Crore" if lp_ub  else "—"
        lp_active   = sum(1 for p in projects if (p.get("project_status") or "").lower() == "active")
        lp_count_fmt = f"{lp_active} Active" if lp_active else f"{lp_count} Total"
    except Exception:
        lp_count_fmt = "—"; lp_ov_fmt = "—"; lp_col_fmt = "—"; lp_ub_fmt = "—"

    return f"""<!DOCTYPE html>
<html lang="en"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<link rel="icon" href="{LOGO_URL}" type="image/jpeg">
<title>Sign In | Monarch Surveyors</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',system-ui,Arial,sans-serif;background:{C['dark']};min-height:100vh;
     display:flex;align-items:center;justify-content:center;padding:20px}}
.wrap{{display:flex;width:100%;max-width:960px;min-height:580px;background:#fff;
       border-radius:18px;overflow:hidden;box-shadow:0 32px 80px rgba(0,0,0,.45)}}
.lp{{width:380px;flex-shrink:0;background:{C['dark']};padding:44px 36px;
     display:flex;flex-direction:column;justify-content:space-between}}
.lp-logo-wrap{{width:64px;height:64px;border-radius:14px;background:#fff;padding:8px;
               display:flex;align-items:center;justify-content:center;margin-bottom:20px;
               box-shadow:0 4px 14px rgba(0,0,0,.3)}}
.lp-logo-wrap img{{width:100%;height:100%;object-fit:contain;border-radius:6px}}
.lp-title{{font-size:15px;font-weight:700;color:#fff;line-height:1.4;margin-bottom:6px}}
.lp-eyebrow{{font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;
             color:{C['secondary']};margin-bottom:24px}}
.lp-tag{{font-size:13px;color:rgba(255,255,255,.55);line-height:1.6;margin-bottom:32px}}
.kpi-row{{display:flex;align-items:center;gap:12px;padding:10px 14px;border-radius:8px;
           border:1px solid rgba(255,255,255,.07);background:rgba(255,255,255,.04);margin-bottom:8px}}
.kpi-icon{{width:34px;height:34px;border-radius:7px;display:flex;align-items:center;
            justify-content:center;font-size:15px;flex-shrink:0}}
.ki-gold{{background:rgba(119,90,25,.35);color:{C['secondary']}}}
.ki-blue{{background:rgba(44,62,107,.5);color:#adc7f8}}
.ki-rose{{background:rgba(186,26,26,.25);color:#f7c1c1}}
.kpi-lbl{{font-size:11px;color:rgba(255,255,255,.5);text-transform:uppercase;margin-bottom:2px}}
.kpi-val{{font-size:15px;font-weight:700;color:#fff}}
.lp-footer{{font-size:11px;color:rgba(255,255,255,.4);line-height:1.5}}
.rp{{flex:1;padding:48px 52px;display:flex;flex-direction:column;justify-content:center}}
.eyebrow{{font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:{C['secondary']};margin-bottom:8px}}
.title{{font-size:30px;font-weight:800;color:{C['dark']};margin-bottom:6px}}
.sub{{font-size:14px;color:#6b7280;margin-bottom:28px}}
.alert-err{{background:#fff1f0;color:{C['error']};border:1px solid #ffb4ab;
             padding:12px 16px;border-radius:9px;font-size:13px;font-weight:600;margin-bottom:20px}}
.rf{{display:flex;flex-direction:column;gap:4px;margin-bottom:16px}}
.rf label{{font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['on_var']}}}
.rf-wrap{{position:relative}}
.rf-ico{{position:absolute;left:13px;top:50%;transform:translateY(-50%);font-size:16px;color:#9ca3af;pointer-events:none}}
.rf input{{width:100%;height:46px;border:1.5px solid {C['outline']};border-radius:9px;
            padding:0 14px 0 42px;font-size:14px;color:{C['on_surface']};background:#fff;outline:none}}
.rf input:focus{{border-color:{C['dark']};box-shadow:0 0 0 4px rgba(0,14,36,.07)}}
.submit-btn{{width:100%;height:48px;background:{C['dark']};color:#fff;border:none;border-radius:9px;
              font-size:13px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;cursor:pointer;margin-top:8px}}
.submit-btn:hover{{background:{C['mid']}}}
.hint{{margin-top:20px;padding:12px 16px;background:{C['surface_low']};border-radius:8px;
        border:1px solid {C['outline']};font-size:12px;color:#6b7280;display:flex;align-items:center;gap:8px}}
.hint code{{background:rgba(0,14,36,.08);color:{C['dark']};padding:1px 6px;border-radius:4px;font-size:11.5px;font-weight:700}}
.footer{{margin-top:20px;font-size:11px;color:#9ca3af;text-align:center}}
@media(max-width:780px){{.lp{{display:none}}.wrap{{max-width:480px}}.rp{{padding:40px 32px}}}}
</style></head><body>

<div class="wrap">

  <!-- LEFT — branding panel -->
  <section class="lp">
    <div>
      <div class="lp-logo-wrap">
        <img src="{LOGO_URL}" alt="Monarch logo">
      </div>
      <p class="lp-title">Monarch Surveyors &amp;<br>Engineering Consultants Ltd.</p>
      <p class="lp-eyebrow">Financial Monitoring System</p>
      <p class="lp-tag">Real-time visibility into project order values, expenditure, collections, and portfolio risk — all from a single live dashboard.</p>

      <div class="kpi-row">
        <div class="kpi-icon ki-blue">📁</div>
        <div>
          <div class="kpi-lbl">Project Portfolio</div>
          <div class="kpi-val">{lp_count_fmt}</div>
        </div>
      </div>
      <div class="kpi-row">
        <div class="kpi-icon ki-gold">💰</div>
        <div>
          <div class="kpi-lbl">Order Value</div>
          <div class="kpi-val">{lp_ov_fmt}</div>
        </div>
      </div>
      <div class="kpi-row">
        <div class="kpi-icon ki-rose">📊</div>
        <div>
          <div class="kpi-lbl">Collection</div>
          <div class="kpi-val">{lp_col_fmt}</div>
        </div>
      </div>
    </div>
    <div class="lp-footer">Secure login &middot; Session-based authentication</div>
  </section>

  <!-- RIGHT — sign-in form -->
  <section class="rp">
    <p class="eyebrow" id="lp-eyebrow">Financial Monitoring System</p>
    <h1 class="title" id="lp-heading">Welcome back</h1>
    <p class="sub" id="lp-sub">Sign in to access your project dashboard and reports.</p>
    {alert_html}
    <form method="post" action="/login">
      <input type="hidden" name="next" id="lf-next" value="{next_path}">
      <div class="rf">
        <label for="lf-user">Username</label>
        <div class="rf-wrap">
          <span class="rf-ico">👤</span>
          <input type="text" id="lf-user" name="username" placeholder="e.g. admin"
                 autocomplete="username" required>
        </div>
      </div>
      <div class="rf">
        <label for="lf-pass">Password</label>
        <div class="rf-wrap">
          <span class="rf-ico">🔒</span>
          <input type="password" id="lf-pass" name="password"
                 placeholder="••••••••" autocomplete="current-password" required>
        </div>
      </div>
      <button type="submit" class="submit-btn" id="lf-submit">Sign in to Dashboard →</button>
    </form>
    <p style="text-align:center;margin-top:16px">
      <a href="/reset_password" style="font-size:12px;font-weight:600;color:{C['secondary']};text-decoration:none">
        Forgot your password? Reset it here →
      </a>
    </p>
    <p class="footer">© {year} Monarch Surveyors &amp; Engineering Consultants Ltd.</p>
  </section>

</div>
<script>
document.getElementById('lf-user').focus();

var PORTAL_COPY = {{
  main: {{
    eyebrow: 'Financial Monitoring System',
    heading: 'Welcome back',
    sub: 'Sign in to access your project dashboard and reports.',
    submit: 'Sign in to Dashboard →',
    next: '/dashboard'
  }},
  payments: {{
    eyebrow: 'Payments Portal',
    heading: 'Welcome to Payments',
    sub: 'Sign in to manage weekly payment bookings and approvals.',
    submit: 'Sign in to Payments →',
    next: '/payments_dashboard'
  }}
}};

function setPortal(p) {{
  var copy = PORTAL_COPY[p] || PORTAL_COPY.main;
  document.getElementById('lp-eyebrow').textContent = copy.eyebrow;
  document.getElementById('lp-heading').textContent = copy.heading;
  document.getElementById('lp-sub').textContent      = copy.sub;
  document.getElementById('lf-submit').textContent   = copy.submit;
  document.getElementById('lf-next').value           = copy.next;
}}

// Pre-select "Payments" if we were bounced here while trying to reach a payments page
(function() {{
  var next = document.getElementById('lf-next').value || '';
  var startPortal = (next.indexOf('/payment') === 0 || next.indexOf('/weekly_booking') === 0) ? 'payments' : 'main';
  setPortal(startPortal);
}})();
</script>
</body></html>"""
def reset_password_page(alert_html: str = "", prefill_username: str = "") -> str:
    year = datetime.now().year
    return f"""<!DOCTYPE html>
<html lang="en"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<link rel="icon" href="{LOGO_URL}" type="image/jpeg">
<title>Reset Password | Monarch Surveyors</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',system-ui,Arial,sans-serif;background:{C['dark']};min-height:100vh;
     display:flex;align-items:center;justify-content:center;padding:20px}}
.wrap{{width:100%;max-width:440px;background:#fff;border-radius:18px;overflow:hidden;
       box-shadow:0 32px 80px rgba(0,0,0,.45);padding:44px 40px}}
.rp-logo-wrap{{width:56px;height:56px;border-radius:12px;background:{C['dark']};padding:7px;
               display:flex;align-items:center;justify-content:center;margin-bottom:20px}}
.rp-logo-wrap img{{width:100%;height:100%;object-fit:contain;border-radius:6px}}
.eyebrow{{font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:{C['secondary']};margin-bottom:8px}}
.title{{font-size:26px;font-weight:800;color:{C['dark']};margin-bottom:6px}}
.sub{{font-size:13px;color:#6b7280;margin-bottom:24px;line-height:1.5}}
.alert-err{{background:#fff1f0;color:{C['error']};border:1px solid #ffb4ab;
             padding:12px 16px;border-radius:9px;font-size:13px;font-weight:600;margin-bottom:20px}}
.alert-ok{{background:#e8f5e9;color:#1b5e20;border:1px solid #a5d6a7;
            padding:12px 16px;border-radius:9px;font-size:13px;font-weight:600;margin-bottom:20px}}
.rf{{display:flex;flex-direction:column;gap:4px;margin-bottom:16px}}
.rf label{{font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['on_var']}}}
.rf-wrap{{position:relative}}
.rf-ico{{position:absolute;left:13px;top:50%;transform:translateY(-50%);font-size:16px;color:#9ca3af;pointer-events:none}}
.rf input{{width:100%;height:46px;border:1.5px solid {C['outline']};border-radius:9px;
            padding:0 14px 0 42px;font-size:14px;color:{C['on_surface']};background:#fff;outline:none}}
.rf input:focus{{border-color:{C['dark']};box-shadow:0 0 0 4px rgba(0,14,36,.07)}}
.submit-btn{{width:100%;height:48px;background:{C['dark']};color:#fff;border:none;border-radius:9px;
              font-size:13px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;cursor:pointer;margin-top:8px}}
.submit-btn:hover{{background:{C['mid']}}}
.back-link{{display:block;text-align:center;margin-top:18px;font-size:12px;font-weight:600;color:{C['on_var']};text-decoration:none}}
.back-link:hover{{color:{C['dark']}}}
.footer{{margin-top:24px;font-size:11px;color:#9ca3af;text-align:center}}
</style></head><body>
<div class="wrap">
  <div class="rp-logo-wrap">
    <img src="{LOGO_URL}" alt="Monarch logo">
  </div>
  <p class="eyebrow">Account Recovery</p>
  <h1 class="title">Reset Password</h1>
  <p class="sub">Enter your username and choose a new password. Contact your administrator if you no longer have access to this account.</p>
  {alert_html}
  <form method="post" action="/reset_password">
    <div class="rf">
      <label for="rp-user">Username</label>
      <div class="rf-wrap">
        <span class="rf-ico">👤</span>
        <input type="text" id="rp-user" name="username" placeholder="e.g. admin"
               value="{prefill_username}" autocomplete="username" required>
      </div>
    </div>
    <div class="rf">
      <label for="rp-new">New Password</label>
      <div class="rf-wrap">
        <span class="rf-ico">🔒</span>
        <input type="password" id="rp-new" name="new_password" placeholder="At least 6 characters"
               autocomplete="new-password" required minlength="6">
      </div>
    </div>
    <div class="rf">
      <label for="rp-conf">Confirm New Password</label>
      <div class="rf-wrap">
        <span class="rf-ico">🔒</span>
        <input type="password" id="rp-conf" name="confirm_password" placeholder="Re-enter password"
               autocomplete="new-password" required minlength="6">
      </div>
    </div>
    <button type="submit" class="submit-btn">Update Password →</button>
  </form>
  <a href="/login" class="back-link">← Back to Sign In</a>
  <p class="footer">© {year} Monarch Surveyors &amp; Engineering Consultants Ltd.</p>
</div>
</body></html>"""

# ═══════════════════════════════════════════════════════════════
#  PAGE 1 — DASHBOARD
# ═══════════════════════════════════════════════════════════════
from datetime import datetime
# ── LOB palette (must stay in same order as get_lob_list()) ──────────────────
_LOB_COLORS = [
    "#1d4ed8",  # blue       — DPR
    "#f97316",  # orange     — DPR-Highways
    "#16a34a",  # green      — ELU preparation
    "#9333ea",  # purple     — Geospatial
    "#0891b2",  # cyan       — Geotechnical
    "#dc2626",  # red        — Land Acquisition
    "#ca8a04",  # yellow     — Locals
    "#be185d",  # pink       — Project Management
    "#0f766e",  # teal       — RAILWAY PROJECT
    "#65a30d",  # lime       — Railways
    "#ea580c",  # deep orange— Railway Tamping
]

from datetime import datetime

_LOB_COLORS = [
    "#1d4ed8",  # blue       — DPR
    "#f97316",  # orange     — DPR-Highways
    "#16a34a",  # green      — ELU preparation
    "#9333ea",  # purple     — Geospatial
    "#0891b2",  # cyan       — Geotechnical
    "#dc2626",  # red        — Land Acquisition
    "#ca8a04",  # yellow     — Locals
    "#be185d",  # pink       — Project Management
    "#0f766e",  # teal       — RAILWAY PROJECT
    "#65a30d",  # lime       — Railways
    "#ea580c",  # deep orange— Railway Tamping
]

def export_dashboard_pdf_bytes() -> bytes:
    projects = fetch_all_projects()
    lob_list = get_lob_list() or []
    def _lob(p):
        val = (p.get("business_line") or p.get("hier_business_line") or "Other").strip()
        return normalize_lob_name(val)

    agg = {}
    for p in projects:
        lb = _lob(p)
        d = agg.setdefault(lb, {"count":0,"ov":0.0,"exp":0.0,"col":0.0})
        d["count"] += 1
        d["ov"]    += _rev_ov(p)
        d["exp"]   += _n(p.get("itd_cost"))
        d["col"]   += _n(p.get("collection_till_date"))

    ordered = []
    seen = set()
    for lb in lob_list:
        norm = normalize_lob_name(lb)
        if norm and norm.lower() not in seen:
            seen.add(norm.lower())
            ordered.append(norm)
    for lb in agg.keys():
        if lb and lb.lower() not in seen:
            seen.add(lb.lower())
            ordered.append(lb)

    # Match the on-screen Dashboard: exclude LOBs with zero Order Value and zero Expenditure.
    ordered = [lb for lb in ordered if agg.get(lb, {}).get("ov", 0) > 0 or agg.get(lb, {}).get("exp", 0) > 0]
    rows = [[lb, agg.get(lb,{}).get("count",0),
              f"{agg.get(lb,{}).get('ov',0):,.0f}",
              f"{agg.get(lb,{}).get('exp',0):,.0f}",
              f"{agg.get(lb,{}).get('col',0):,.0f}"] for lb in ordered]

    total_ov = sum(_rev_ov(p) for p in projects)
    return _build_table_pdf(
        "Executive Performance Dashboard",
        f"LOB performance summary — {len(projects)} projects",
        ["Line of Business","Projects","Order Value (₹)","Expenditure (₹)","Collection (₹)"],
        rows, footer_label="Total Order Value", footer_value=total_ov)

def dashboard(username: str = "Admin", is_admin: bool = False) -> str:
    active_only = bool(getattr(_req_ctx, "active_only", False))
    projects   = fetch_all_projects()   # already Active-only if the global toggle is on
    total_proj = len(projects)
    lob_list   = get_visible_lob_list() or []

    # ── Aggregate totals ──────────────────────────────────────────────────────
    total_ov     = sum(_rev_ov(p)                         for p in projects)

    # Total Billed / Total Collected are summed over PARENT projects only,
    # because a parent's own "billed_amount" / "collection_till_date" is
    # normally already the rolled-up figure for all its sub-projects --
    # summing both would double-count. However a few projects (e.g. co_no
    # "366 A") have their billing/collection entered directly on the
    # sub-projects instead, leaving the parent's own fields at 0 -- which
    # used to make that money disappear from these two KPIs entirely (it
    # was excluded as a "child" row, but never rolled up to its parent
    # either). _eff_billed_col() closes that gap: it only falls back to
    # summing a parent's children when the parent's OWN field is zero and
    # its children actually carry a non-zero amount. Any parent that
    # already has a real value of its own (the normal, common case) is
    # completely untouched by this fallback, so nothing else changes.
    _parents_only  = _rep_parents_only(projects)
    _all_children  = _rep_child_projects(projects)

    def _eff_billed_col(parent):
        own_billed = _n(parent.get("billed_amount"))
        own_col    = _n(parent.get("collection_till_date"))
        if own_billed == 0 or own_col == 0:
            kids = _rep_kids_of(parent, _all_children)
            if kids:
                if own_billed == 0:
                    kid_billed = sum(_n(c.get("billed_amount")) for c in kids)
                    if kid_billed > 0:
                        own_billed = kid_billed
                if own_col == 0:
                    kid_col = sum(_n(c.get("collection_till_date")) for c in kids)
                    if kid_col > 0:
                        own_col = kid_col
        return own_billed, own_col

    _eff_by_id   = {id(p): _eff_billed_col(p) for p in _parents_only}
    total_billed = sum(v[0] for v in _eff_by_id.values())
    total_col    = sum(v[1] for v in _eff_by_id.values())
    total_exp    = sum(_n(p.get("itd_cost"))             for p in projects)
    total_rev_ov = sum(_n(p.get("revised_order_value") or (_n(p.get("order_value")) + _n(p.get("variation_amt")))) for p in projects)
    # Matches Consolidated Reports' "Unbilled Amount": Revised Order Value (or
    # Order Value if no revision) minus Billed Amount, summed over parent
    # projects only (so sub-project figures already rolled into a parent
    # aren't double-counted). Uses the same effective billed amount as
    # above so a parent whose billing lives on its children isn't shown as
    # fully unbilled here either.
    unbilled     = sum(
        max(_rep_oi_amt(p) - _eff_by_id[id(p)][0], 0)
        for p in _parents_only
    )
    # Pending Collection: amount already billed but not yet collected
    # (same "Billed but Not Collected" logic used in Consolidated Reports),
    # summed over parent projects only so sub-project figures already rolled
    # into a parent aren't double-counted.
    total_pending_col = max(total_billed - total_col, 0)
    avg_comp     = (
        sum(_n(p.get("physical_completion")) for p in projects) / total_proj
        if total_proj else 0
    )
    col_pct = f"{total_col / total_ov * 100:.1f}%" if total_ov else "—"
    comp_w  = f"{avg_comp:.1f}%"
    
    # ── Per-LOB aggregates ────────────────────────────────────────────────────
    def _lob(p):
        val = (p.get("business_line") or p.get("hier_business_line") or "Other").strip()
        return normalize_lob_name(val)

    lob_data = {}
    for p in projects:
        lb = _lob(p)
        if lb not in lob_data:
            lob_data[lb] = {"count": 0, "ov": 0.0, "exp": 0.0,
                            "col": 0.0, "billed": 0.0}
        lob_data[lb]["count"]  += 1
        lob_data[lb]["ov"]     += _rev_ov(p)
        lob_data[lb]["exp"]    += _n(p.get("itd_cost"))
        lob_data[lb]["col"]    += _n(p.get("collection_till_date"))
        lob_data[lb]["billed"] += _n(p.get("billed_amount"))

    ordered_lobs = []
    seen_lobs = set()
    for lb in (lob_list or []):
        norm = normalize_lob_name(lb)
        if norm and norm.lower() not in seen_lobs:
            seen_lobs.add(norm.lower())
            ordered_lobs.append(norm)
    for lb in lob_data.keys():
        if lb and lb.lower() not in seen_lobs:
            seen_lobs.add(lb.lower())
            ordered_lobs.append(lb)

    # ── Exclude LOBs with zero Order Value and zero Expenditure from the Dashboard entirely
    # (chart, legend, and table all derive from ordered_lobs, so filtering
    # here keeps them consistent automatically). ──
    ordered_lobs = [lb for lb in ordered_lobs if lob_data.get(lb, {}).get("ov", 0) > 0 or lob_data.get(lb, {}).get("exp", 0) > 0]

    # ── Formatting helpers ────────────────────────────────────────────────────
    def _fmt(v: float) -> str:
        if v >= 1e7:
            return f"&#8377; {v / 1e7:.2f} Cr"
        if v >= 1e5:
            return f"&#8377; {v / 1e5:.1f} L"
        return f"&#8377; {v:,.0f}"

    def _fmt_short(v: float) -> str:
        if v >= 1e7:
            return f"&#8377; {v / 1e7:.1f} Cr"
        return f"&#8377; {v:,.0f}"

    # ── Chart.js JS data arrays ───────────────────────────────────────────────
    def _js_arr(vals):
        return "[" + ",".join(str(round(v, 2)) for v in vals) + "]"

    lob_colors_js = "[" + ",".join(f'"{_LOB_COLORS[i % len(_LOB_COLORS)]}"'
                                   for i in range(len(ordered_lobs))) + "]"
    lob_labels_js = "[" + ",".join(f'"{lb}"' for lb in ordered_lobs) + "]"

    counts_js = _js_arr([lob_data.get(lb, {}).get("count", 0) for lb in ordered_lobs])
    ov_js     = _js_arr([lob_data.get(lb, {}).get("ov",    0) for lb in ordered_lobs])
    exp_js    = _js_arr([lob_data.get(lb, {}).get("exp",   0) for lb in ordered_lobs])
    col_js    = _js_arr([lob_data.get(lb, {}).get("col",   0) for lb in ordered_lobs])
    wip_js    = _js_arr([
        max(lob_data.get(lb, {}).get("billed", 0) - lob_data.get(lb, {}).get("col", 0), 0)
        for lb in ordered_lobs
    ])

    # ── KPI cards ─────────────────────────────────────────────────────────────
    def _kcard(label, val, sub_html, href=""):
        tag, extra_attrs, extra_style = (
            ("a", f' href="{href}"', "text-decoration:none;cursor:pointer;transition:box-shadow .15s;")
            if href else ("div", "", "")
        )
        return f"""
        <{tag}{extra_attrs} style="display:block;background:#fff;border-radius:10px;padding:16px 18px;
                    border:0.5px solid {C['outline']};{extra_style}">
          <p style="font-size:10px;font-weight:700;letter-spacing:.07em;
                    color:#9ca3af;text-transform:uppercase;margin-bottom:8px">{label}</p>
          <p style="font-size:24px;font-weight:800;color:{C['dark']};line-height:1">{val}</p>
          {sub_html}
        </{tag}>"""

    # ── Real month-over-month project growth (based on created_at) ──
    now               = datetime.now()
    this_month_start  = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    prev_month_end    = this_month_start - timedelta(seconds=1)
    prev_month_start  = prev_month_end.replace(day=1, hour=0, minute=0, second=0, microsecond=0)

    def _created_dt(p):
        try:
            return datetime.fromisoformat((p.get("created_at") or "")[:19])
        except Exception:
            return None

    count_prev_month_end = sum(
        1 for p in projects
        if (c := _created_dt(p)) and c < this_month_start
    )
    count_now = total_proj

    if count_prev_month_end > 0:
        growth_pct = (count_now - count_prev_month_end) / count_prev_month_end * 100
        arrow      = "&#8593;" if growth_pct >= 0 else "&#8595;"
        color      = "#16a34a" if growth_pct >= 0 else C["error"]
        sign       = "+" if growth_pct >= 0 else ""
        growth_html = (
            f'<p style="font-size:11px;color:{color};font-weight:600;margin-top:6px">'
            f'{arrow} {sign}{growth_pct:.1f}% vs Prev. Month</p>'
        )
    elif count_now > 0:
        growth_html = (
            '<p style="font-size:11px;color:#16a34a;font-weight:600;margin-top:6px">'
            '&#8593; New this month</p>'
        )
    else:
        growth_html = (
            '<p style="font-size:11px;color:#9ca3af;margin-top:6px">No data yet</p>'
        )

    cards_html = (
        _kcard("Total Project Inventory", str(total_proj), growth_html) +
        _kcard("Aggregate Collection", _fmt_short(total_col),
               f'<p style="font-size:11px;color:#6b7280;margin-top:6px">'
               f'{col_pct} Realization Rate</p>') +
        
        _kcard("Avg. Completion", comp_w,
               f'<div style="background:#e5e7eb;border-radius:99px;height:7px;'
               f'margin-top:10px;overflow:hidden">'
               f'<div style="height:100%;width:{comp_w};background:{C["dark"]};'
               f'border-radius:99px"></div></div>')
    )

    # ── Second KPI row: portfolio-wide financial totals ─────────────────────
    cards2_html = (
        _kcard("Total Order Value", _fmt_short(total_ov),
               '<p style="font-size:11px;color:#6b7280;margin-top:6px">Revised Order Value, All Projects</p>') +
        _kcard("Total Expenditure", _fmt_short(total_exp),
               '<p style="font-size:11px;color:#6b7280;margin-top:6px">Inception-to-Date Cost</p>') +
        _kcard("Total Billed", _fmt_short(total_billed),
               '<p style="font-size:11px;color:#6b7280;margin-top:6px">Up to date Billed Amount</p>') +
        _kcard("Total Unbilled", _fmt_short(unbilled),
               ('<p style="font-size:11px;color:#dc2626;font-weight:600;margin-top:6px">'
                '&#9888; Attention Required</p>'
                if unbilled > 0 else
                '<p style="font-size:11px;color:#16a34a;font-weight:600;margin-top:6px">'
                '&#10003; Clear</p>')) +
        _kcard("Total Collected", _fmt_short(total_col),
               f'<p style="font-size:11px;color:#6b7280;margin-top:6px">{col_pct} Realization Rate</p>') +
        _kcard("Total Pending Collection", _fmt_short(total_pending_col),
               ('<p style="font-size:11px;color:#dc2626;font-weight:600;margin-top:6px">'
                '&#9888; Billed, Not Yet Collected</p>'
                if total_pending_col > 0 else
                '<p style="font-size:11px;color:#16a34a;font-weight:600;margin-top:6px">'
                '&#10003; Clear</p>'))
    )

    # ── Global totals bar ─────────────────────────────────────────────────────
    glob_stats = [
        ("Projects",     f"{total_proj} Units", "#fff"),
        ("Order Value",  _fmt_short(total_ov),  "#fff"),
        ("Expenditure",  _fmt_short(total_exp), "#fff"),
        ("Collection",   _fmt_short(total_col), "#fff"),
        ("Unbilled WIP", _fmt_short(unbilled),  "#f87171"),
    ]
    glob_html = "".join(
        f'<div style="text-align:center;padding:0 16px;'
        f'{"border-right:none" if i == len(glob_stats) - 1 else "border-right:1px solid #1e3a6a"}">'
        f'<div style="font-size:9px;color:#64748b;text-transform:uppercase;'
        f'letter-spacing:.06em;margin-bottom:2px">{lbl}</div>'
        f'<div style="font-size:12px;font-weight:700;color:{clr}">{val}</div>'
        f'</div>'
        for i, (lbl, val, clr) in enumerate(glob_stats)
    )

    # ── LOB table rows ────────────────────────────────────────────────────────
    lob_rows = ""
    for lb in ordered_lobs:
        d   = lob_data.get(lb, {})
        ov  = d.get("ov",    0)
        col = d.get("col",   0)
        exp = d.get("exp",   0)
        cnt = d.get("count", 0)
        pct = f"{col / ov * 100:.1f}%" if ov else "—"
        lob_rows += f"""
        <tr style="border-bottom:0.5px solid #f1f3f7">
          <td style="padding:9px 14px;font-weight:600;color:{C['dark']}">{lb}</td>
          <td style="padding:9px 14px;text-align:right;font-family:'Courier New',monospace">{cnt}</td>
          <td style="padding:9px 14px;text-align:right;font-family:'Courier New',monospace">{_fmt(ov)}</td>
          <td style="padding:9px 14px;text-align:right;font-family:'Courier New',monospace">{_fmt(exp)}</td>
          <td style="padding:9px 14px;text-align:right;font-family:'Courier New',monospace">{_fmt(col)}</td>
          <td style="padding:9px 14px;text-align:right;font-weight:700;color:{C['dark']}">{pct}</td>
        </tr>"""

    if not lob_rows:
        lob_rows = (f'<tr><td colspan="6" style="padding:20px;text-align:center;color:#aaa">'
                    f'No projects yet. <a href="/data_entry">Add via Data Entry &#8594;</a></td></tr>')

    # ── LOB legend ────────────────────────────────────────────────────────────
    legend_html = "".join(
        f'<span style="display:flex;align-items:center;gap:5px;font-size:10px;color:#6b7280">'
        f'<span style="width:10px;height:10px;border-radius:2px;'
        f'background:{_LOB_COLORS[i % len(_LOB_COLORS)]}"></span>{lb}</span>'
        for i, lb in enumerate(ordered_lobs)
    )

    # ── Chart.js + custom fixed tooltip (never clipped by any parent) ─────────
    chart_script = f"""
<style>
#ch-tooltip {{
  position: fixed;
  background: rgba(15,27,61,0.93);
  color: #fff;
  padding: 8px 12px;
  border-radius: 7px;
  font-size: 12px;
  font-family: sans-serif;
  pointer-events: none;
  z-index: 99999;
  display: none;
  max-width: 240px;
  line-height: 1.5;
  box-shadow: 0 4px 16px rgba(0,0,0,0.35);
}}
#ch-tooltip .tt-title {{
  font-weight: 700;
  font-size: 13px;
  margin-bottom: 5px;
  padding-bottom: 5px;
  border-bottom: 1px solid rgba(255,255,255,0.18);
}}
#ch-tooltip .tt-row {{
  display: flex;
  align-items: center;
  gap: 7px;
}}
#ch-tooltip .tt-swatch {{
  width: 10px;
  height: 10px;
  border-radius: 2px;
  flex-shrink: 0;
}}
#ch-tooltip .tt-val {{
  font-weight: 600;
  font-size: 13px;
}}
</style>
<div id="ch-tooltip">
  <div class="tt-title" id="tt-title"></div>
  <div class="tt-row">
    <div class="tt-swatch" id="tt-swatch"></div>
    <div class="tt-val" id="tt-body"></div>
  </div>
</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<script>
(function () {{
  var ttBox    = document.getElementById('ch-tooltip');
  var ttTitle  = document.getElementById('tt-title');
  var ttBody   = document.getElementById('tt-body');
  var ttSwatch = document.getElementById('tt-swatch');

  function fmtVal(v, isCount) {{
    if (isCount) return v + ' units';
    if (v >= 10000000) return '\u20b9 ' + (v / 10000000).toFixed(2) + ' Cr';
    if (v >= 100000)   return '\u20b9 ' + (v / 100000).toFixed(1) + ' L';
    return '\u20b9 ' + v.toFixed(0).replace(/\\B(?=(\\d{{3}})+(?!\\d))/g, ',');
  }}

  function customTooltip(context, isCount) {{
    var model = context.tooltip;
    if (model.opacity === 0) {{
      ttBox.style.display = 'none';
      return;
    }}
    if (model.dataPoints && model.dataPoints.length) {{
      var dp = model.dataPoints[0];
      ttTitle.textContent          = dp.label;
      ttBody.textContent           = fmtVal(dp.raw, isCount);
      ttSwatch.style.background    = model.labelColors[0].backgroundColor;
    }}
    var rect = context.chart.canvas.getBoundingClientRect();
    var x    = rect.left + model.caretX;
    var y    = rect.top  + model.caretY;
    var ttW  = ttBox.offsetWidth  || 180;
    var ttH  = ttBox.offsetHeight || 60;
    var winW = window.innerWidth;
    var left = x + 14;
    var top  = y - ttH - 10;
    if (left + ttW > winW - 10) {{ left = x - ttW - 14; }}
    if (top < rect.top)         {{ top  = y + 14; }}
    ttBox.style.left    = left + 'px';
    ttBox.style.top     = top  + 'px';
    ttBox.style.display = 'block';
  }}

  var colors = {lob_colors_js};
  var labels = {lob_labels_js};
  var sets = [
    {{ id: 'ch0', data: {counts_js}, isCount: true  }},
    {{ id: 'ch1', data: {ov_js},     isCount: false }},
    {{ id: 'ch2', data: {exp_js},    isCount: false }},
    {{ id: 'ch3', data: {col_js},    isCount: false }},
    {{ id: 'ch4', data: {wip_js},    isCount: false }}
  ];

  sets.forEach(function (s) {{
    var el = document.getElementById(s.id);
    if (!el) return;
    var isCount = s.isCount;
    new Chart(el, {{
      type: 'doughnut',
      data: {{
        labels: labels,
        datasets: [{{
          data: s.data,
          backgroundColor: colors,
          borderWidth: 0,
          hoverOffset: 6
        }}]
      }},
      options: {{
        responsive: false,
        cutout: '62%',
        plugins: {{
          legend: {{ display: false }},
          tooltip: {{
            enabled: false,
            external: function (ctx) {{ customTooltip(ctx, isCount); }}
          }}
        }}
      }}
    }});
  }});

  document.addEventListener('mouseleave', function () {{
    ttBox.style.display = 'none';
  }});
}})();
</script>"""

    # ── Footer ────────────────────────────────────────────────────────────────
    now     = datetime.now().strftime("%d %b %Y, %I:%M %p IST")
    pending = fetch_summary_stats().get("pending", 0)

    # ── Assemble full page ────────────────────────────────────────────────────
    html  = _html_head("Executive Dashboard") + "<body>"
    html += _sidebar("dashboard", is_admin=is_admin)
    html += _topbar("FINANCIAL MONITORING SYSTEM", username)
    html += f"""
<div class="main"><div class="canvas">

  <!-- Page header -->
  <p style="font-size:11px;color:#9ca3af;margin-bottom:6px">
    REPORTS / <span style="color:#374151;font-weight:500">LOB PERFORMANCE SUMMARY</span>
  </p>
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px">
    <div>
      <h1 style="font-size:22px;font-weight:800;color:{C['dark']}">
        Executive Performance Dashboard
      </h1>
      <p style="font-size:12px;color:#6b7280;margin-top:2px">
        Live data from database &mdash; {total_proj} projects loaded
      </p>
    </div>
    <div style="display:flex;align-items:center;gap:10px">
      <a href="/toggle_active_filter?next={_urlq('/dashboard')}"
         style="display:inline-flex;align-items:center;gap:6px;
                background:{'#16a34a' if active_only else '#fff'};
                color:{'#fff' if active_only else C['dark']};
                border:1px solid {'#16a34a' if active_only else C['outline']};
                padding:10px 20px;border-radius:999px;font-size:13px;font-weight:700;
                text-decoration:none;letter-spacing:.01em;cursor:pointer;white-space:nowrap"
         title="{'Showing Active projects only across the app — click to show all' if active_only else 'Click to show only Active projects across the whole app'}">
        {'✅ Active Only' if active_only else '◻ All Projects'}
      </a>
      <a href="/export_excel/full" download
         style="display:inline-flex;align-items:center;gap:8px;background:{C['primary']};
                color:#ffffff;padding:11px 22px;border-radius:999px;font-size:13.5px;
                font-weight:700;text-decoration:none;letter-spacing:.01em;border:none;
                box-shadow:0 2px 6px rgba(0,14,36,.25);cursor:pointer;white-space:nowrap">
        ⬇ Export Excel
      </a>
    </div>
  </div>

  <!-- KPI cards -->
  <div style="display:grid;grid-template-columns:repeat(4,1fr);
              gap:14px;margin-bottom:14px">
    {cards_html}
  </div>

  <!-- KPI cards — portfolio financial totals -->
  <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));
              gap:14px;margin-bottom:18px">
    {cards2_html}
  </div>

  <!-- Global totals bar -->
  <div style="background:{C['dark']};border-radius:10px 10px 0 0;
              padding:13px 20px;display:flex;align-items:center;flex-wrap:wrap;
              gap:10px;overflow-x:auto">
    <span style="font-size:10px;font-weight:700;letter-spacing:.07em;
                 color:#64748b;text-transform:uppercase;margin-right:auto;
                 white-space:nowrap">
      Global Portfolio Totals
    </span>
    {glob_html}
  </div>
  <!-- LOB breakdown table -->
  <div style="background:#fff;border-radius:0 0 10px 10px;
              border:0.5px solid {C['outline']};border-top:none;margin-bottom:18px">
    <div style="padding:12px 16px;border-bottom:0.5px solid {C['outline']};
                display:flex;justify-content:space-between;align-items:center;
                background:{C['surface_low']}">
      <span style="font-size:13px;font-weight:700;color:{C['dark']}">
        LOB Performance Breakdown
      </span>
      <a href="/data_entry" style="background:{C['dark']};color:#fff;
         border-radius:8px;padding:5px 10px;font-size:11px;text-decoration:none">
        + Add Project
      </a>
    </div>
    <div style="overflow-x:auto">
      <table style="width:100%;min-width:800px;border-collapse:collapse;font-size:12px">
        <thead>
          <tr style="background:{C['surface_low']};border-bottom:0.5px solid {C['outline']}">
            <th style="padding:9px 14px;text-align:left;font-size:10px;font-weight:700;
                       letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']}">Line of Business</th>
            <th style="padding:9px 14px;text-align:right;font-size:10px;font-weight:700;
                       letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']}">Projects</th>
            <th style="padding:9px 14px;text-align:right;font-size:10px;font-weight:700;
                       letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']}">Order Value (&#8377;)</th>
            <th style="padding:9px 14px;text-align:right;font-size:10px;font-weight:700;
                       letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']}">Expenditure (&#8377;)</th>
            <th style="padding:9px 14px;text-align:right;font-size:10px;font-weight:700;
                       letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']}">Collection (&#8377;)</th>
            <th style="padding:9px 14px;text-align:right;font-size:10px;font-weight:700;
                       letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']}">Realization %</th>
          </tr>
        </thead>
        <tbody>{lob_rows}</tbody>
      </table>
    </div>
  </div>

  <!-- 5 doughnut charts -->
  <div style="background:#fff;border-radius:10px;border:0.5px solid {C['outline']};
              margin-bottom:18px;overflow-x:auto">
    <div style="display:grid;grid-template-columns:repeat(4,minmax(150px,1fr));
                gap:0;padding:20px 8px;min-width:760px">
      <div style="display:flex;flex-direction:column;align-items:center;
                  border-right:0.5px solid {C['outline']};padding:0 12px">
        <p style="font-size:9px;font-weight:700;letter-spacing:.06em;
                  text-transform:uppercase;color:#9ca3af;margin-bottom:8px">Ongoing Projects</p>
        <canvas id="ch0" width="130" height="130"
                role="img" aria-label="Doughnut chart of ongoing projects by LOB"></canvas>
        <p style="font-size:12px;font-weight:800;color:{C['dark']};margin-top:8px">N = {total_proj}</p>
        <p style="font-size:9px;color:#9ca3af;margin-top:1px">Total unit count</p>
      </div>

      <div style="display:flex;flex-direction:column;align-items:center;
                  border-right:0.5px solid {C['outline']};padding:0 12px">
        <p style="font-size:9px;font-weight:700;letter-spacing:.06em;
                  text-transform:uppercase;color:#9ca3af;margin-bottom:8px">Order Value</p>
        <canvas id="ch1" width="130" height="130"
                role="img" aria-label="Doughnut chart of order value by LOB"></canvas>
        <p style="font-size:12px;font-weight:800;color:{C['dark']};margin-top:8px">{_fmt_short(total_ov)}</p>
        <p style="font-size:9px;color:#9ca3af;margin-top:1px">Projected revenue</p>
      </div>

      <div style="display:flex;flex-direction:column;align-items:center;
                  border-right:0.5px solid {C['outline']};padding:0 12px">
        <p style="font-size:9px;font-weight:700;letter-spacing:.06em;
                  text-transform:uppercase;color:#9ca3af;margin-bottom:8px">Expenditure</p>
        <canvas id="ch2" width="130" height="130"
                role="img" aria-label="Doughnut chart of expenditure by LOB"></canvas>
        <p style="font-size:12px;font-weight:800;color:{C['dark']};margin-top:8px">{_fmt_short(total_exp)}</p>
        <p style="font-size:9px;color:#9ca3af;margin-top:1px">Cost incurred</p>
      </div>

      <div style="display:flex;flex-direction:column;align-items:center;
                  border-right:0.5px solid {C['outline']};padding:0 12px">
        <p style="font-size:9px;font-weight:700;letter-spacing:.06em;
                  text-transform:uppercase;color:#9ca3af;margin-bottom:8px">Collection</p>
        <canvas id="ch3" width="130" height="130"
                role="img" aria-label="Doughnut chart of collection by LOB"></canvas>
        <p style="font-size:12px;font-weight:800;color:{C['dark']};margin-top:8px">{_fmt_short(total_col)}</p>
        <p style="font-size:9px;color:#9ca3af;margin-top:1px">Realized funds</p>
      </div>

      

    </div>
    <!-- LOB legend -->
    <div style="padding:10px 16px 14px;border-top:0.5px solid {C['outline']};
                display:flex;flex-wrap:wrap;gap:14px;justify-content:center">
      {legend_html}
    </div>
  </div>

  <!-- Footer -->
  <div style="display:flex;justify-content:space-between;align-items:center;
              font-size:10px;color:#9ca3af;padding:0 2px;margin-bottom:16px">
    <span>
      <span style="color:#22c55e;margin-right:3px">&#9679;</span>
      Data Last Synced: {now}
    </span>
    <span>&#10003; Audit Status: Verified (V.2.4) &ndash; Monarch Precision Standard</span>
    <span style="color:#dc2626;font-weight:600">
      &#9888; {pending} Variations Pending Final Approval
    </span>
  </div>

  <p style="font-size:10px;color:#d1d5db;text-align:center;margin-bottom:8px">
    &copy; 2024 Monarch Surveyors and Engineering Consultants Ltd. | Confidential
  </p>

</div></div>
{chart_script}
</body></html>"""

    return html
# ═══════════════════════════════════════════════════════════════
#  PAGE 2 — DATA ENTRY
# ═══════════════════════════════════════════════════════════════

def project_data_entry(alert_html="", prefill=None, username="Admin", status_filter="") -> str:
    p = prefill or {}

    ENTRY_CSS = f"""
<style>
.page-title{{font-size:26px;font-weight:800;color:{C['dark']};margin-bottom:4px}}
.page-sub{{font-size:13px;color:#6b7280;margin-bottom:20px}}
.upload-zone{{border:2px dashed {C['outline']};border-radius:12px;padding:28px;background:#fff;text-align:center;cursor:pointer;margin-bottom:24px;transition:border-color .2s}}
.upload-zone:hover{{border-color:{C['dark']};background:#f0f4ff}}
.card{{background:#fff;border:1px solid {C['outline']};border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.06);margin-bottom:20px}}
.tab-bar{{display:flex;border-bottom:1px solid {C['outline']};background:{C['surface_low']};overflow-x:auto}}
.tab-btn{{flex:1;min-width:160px;padding:13px 16px;border:none;background:transparent;font-size:11px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']};cursor:pointer;border-bottom:2px solid transparent;transition:color .15s}}
.tab-btn:hover,.tab-active{{color:{C['dark']}!important;border-bottom-color:{C['dark']}!important}}
.section{{padding:28px;display:none}}
.sec-title{{font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['dark']};margin-bottom:14px}}
.g2{{display:grid;grid-template-columns:1fr 1fr;gap:16px}}
.g3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px}}
.g12{{display:grid;grid-template-columns:1fr 2fr;gap:16px}}
.g21{{display:grid;grid-template-columns:2fr 1fr;gap:16px}}
.mt{{margin-top:16px}}
.field{{display:flex;flex-direction:column;gap:4px}}
.field label{{font-size:11px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']}}}
.field input,.field select,.field textarea{{height:40px;border:1px solid {C['outline']};border-radius:6px;padding:0 12px;font-size:14px;color:{C['on_surface']};background:#fff;outline:none;transition:border-color .15s;width:100%;box-sizing:border-box}}
.field input:focus,.field select:focus{{border-color:{C['dark']};box-shadow:0 0 0 3px rgba(0,14,36,.07)}}
.cur-wrap{{position:relative}}
.cur-wrap .pre{{position:absolute;left:10px;top:50%;transform:translateY(-50%);font-weight:700;color:{C['on_var']};font-size:13px;pointer-events:none}}
.cur-wrap input{{padding-left:24px}}
.suf-wrap{{position:relative}}
.suf-wrap .suf{{position:absolute;right:10px;top:50%;transform:translateY(-50%);color:{C['on_var']};font-size:13px;pointer-events:none}}
.suf-wrap input{{padding-right:28px}}
.fin-card{{background:{C['surface_low']};border:1px solid {C['outline']};border-radius:8px;padding:16px}}
.info-ban{{display:flex;align-items:flex-start;gap:10px;background:#fffbf0;border:1px solid {C['sec_cont']};border-radius:8px;padding:12px 16px;font-size:13px;color:#5d4201;margin-top:14px}}
.proj-table{{width:100%;border-collapse:collapse;font-size:13px}}
.proj-table th{{padding:9px 14px;font-size:11px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']};text-align:left;background:{C['surface_low']};border-bottom:1px solid {C['outline']}}}
.proj-table td{{padding:11px 14px;border-bottom:1px solid #eee;vertical-align:middle}}
.proj-table tr:last-child td{{border-bottom:none}}
.proj-table tr:hover td{{background:{C['surface_low']}}}
.del-btn{{background:transparent;border:none;cursor:pointer;font-size:16px;color:#ef4444;padding:4px 8px;border-radius:4px}}
.del-btn:hover{{background:#fef2f2}}
.exp-bar{{display:flex;gap:10px;margin-bottom:16px}}
.exp-btn{{background:{C['surface_low']};border:1px solid {C['outline']};color:{C['dark']};padding:7px 16px;border-radius:6px;font-size:12px;font-weight:700;text-decoration:none;letter-spacing:.04em;text-transform:uppercase;cursor:pointer}}
.exp-btn:hover{{background:{C['surface_high']};border-color:{C['dark']}}}

/* ── Global save bar ── */
.global-save-bar{{
  position:sticky;
  bottom:0;
  left:0;
  right:0;
  z-index:200;
  display:flex;
  align-items:center;
  justify-content:space-between;
  padding:12px 28px;
  background:rgba(255,255,255,.97);
  border-top:2px solid {C['dark']};
  box-shadow:0 -4px 20px rgba(0,0,0,.10);
  backdrop-filter:blur(6px);
}}
.gsb-pills{{display:flex;align-items:center;gap:8px}}
.gsb-dirty{{
  display:none;
  align-items:center;
  gap:6px;
  font-size:12px;
  font-weight:700;
  color:#b45309;
  background:#fef3c7;
  border:1px solid #fcd34d;
  border-radius:20px;
  padding:4px 12px;
}}
.gsb-dirty.show{{display:inline-flex}}
.btn-gsave{{
  background:{C['dark']};
  color:#fff;
  border:none;
  padding:11px 32px;
  border-radius:8px;
  font-size:13px;
  font-weight:700;
  text-transform:uppercase;
  letter-spacing:.05em;
  cursor:pointer;
  box-shadow:0 2px 8px rgba(0,0,0,.20);
  display:inline-flex;
  align-items:center;
  gap:8px;
  transition:opacity .15s;
}}
.btn-gsave:disabled{{opacity:.6;cursor:not-allowed}}
.btn-gsave:hover:not(:disabled){{opacity:.88}}

/* ── Step pills (inside save bar) ── */
.step-pill{{
  display:flex;align-items:center;gap:6px;padding:5px 12px;border-radius:20px;
  font-size:10px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;
  background:{C['surface_high']};color:{C['on_var']};border:1px solid {C['outline']};
  transition:all .2s;cursor:pointer;
}}
.step-pill.active{{background:{C['dark']};color:#fff;border-color:{C['dark']};}}
.step-pill.done{{background:#e8f5e9;color:#166534;border-color:#a5d6a7;}}
.step-num{{width:18px;height:18px;border-radius:50%;display:inline-flex;align-items:center;
  justify-content:center;font-size:10px;font-weight:800;background:rgba(255,255,255,.25);}}
.step-pill.active .step-num{{background:rgba(255,255,255,.2);}}
.step-pill.done  .step-num{{background:#a5d6a7;color:#166534;}}

/* ── Collapsible panels ── */
.panel-collapsed{{display:none!important}}
.toggle-panel-btn.open{{background:{C['dark']};color:#fff;border-color:{C['dark']}}}
</style>"""

    # ── Document icon SVG helper ──────────────────────────────────────────────
    GDRIVE_SVG = (
        '<svg width="13" height="13" viewBox="0 0 87.3 78" xmlns="http://www.w3.org/2000/svg">'
        '<path d="m6.6 66.85 3.85 6.65c.8 1.4 1.95 2.5 3.3 3.3l13.75-23.8h-27.5c0 1.55.4 3.1 1.2 4.5z" fill="#0066da"/>'
        '<path d="m43.65 25-13.75-23.8c-1.35.8-2.5 1.9-3.3 3.3l-25.4 44a9.06 9.06 0 0 0 -1.2 4.5h27.5z" fill="#00ac47"/>'
        '<path d="m73.55 76.8c1.35-.8 2.5-1.9 3.3-3.3l1.6-2.75 7.65-13.25c.8-1.4 1.2-2.95 1.2-4.5h-27.502l5.852 11.5z" fill="#ea4335"/>'
        '<path d="m43.65 25 13.75-23.8c-1.35-.8-2.9-1.2-4.5-1.2h-18.5c-1.6 0-3.15.45-4.5 1.2z" fill="#00832d"/>'
        '<path d="m59.8 53h-32.3l-13.75 23.8c1.35.8 2.9 1.2 4.5 1.2h50.8c1.6 0 3.15-.45 4.5-1.2z" fill="#2684fc"/>'
        '<path d="m73.4 26.5-12.7-22c-.8-1.4-1.95-2.5-3.3-3.3l-13.75 23.8 16.15 28h27.45c0-1.55-.4-3.1-1.2-4.5z" fill="#ffba00"/>'
        '</svg>'
    )

    TAB_JS = """<script>
var currentTab = 1, totalTabs = 4;
var _dirty = false;

/* ── dirty-tracking ── */
function markDirty() {
  if (_dirty) return;
  _dirty = true;
  var d = document.getElementById('gsb-dirty');
  if (d) d.classList.add('show');
}
function clearDirty() {
  _dirty = false;
  var d = document.getElementById('gsb-dirty');
  if (d) d.classList.remove('show');
}

document.addEventListener('input',  markDirty);
document.addEventListener('change', markDirty);

/* ── filter parent projects by LOB ── */
var _allParentOpts = null;

function filterParentsByLob(lobVal) {
  var picker = document.getElementById('parent-picker-sel');
  var dl     = document.getElementById('dl_parent_project');
  if (!picker || !dl) return;

  if (!lobVal) {
    if (_allParentOpts) {
      picker.innerHTML = _allParentOpts.picker;
      dl.innerHTML     = _allParentOpts.dl;
      window.PARENT_PROJECT_MAP = _allParentOpts.map;
      syncParentId();
    }
    return;
  }

  fetch('/get_parents_by_lob?lob=' + encodeURIComponent(lobVal))
    .then(function(r) { return r.json(); })
    .then(function(items) {
      if (!_allParentOpts) {
        _allParentOpts = { picker: picker.innerHTML, dl: dl.innerHTML, map: window.PARENT_PROJECT_MAP || {} };
      }
      var pickerHtml = '<option value="">— Select a parent project —</option>';
      var dlHtml     = '';
      var newMap     = {};
      var labels     = [];
      items.forEach(function(it) {
        pickerHtml += '<option value="' + it.label + '">' + it.label + '</option>';
        dlHtml     += '<option value="' + it.label + '">';
        newMap[it.label] = it.co_no;
        newMap[it.name]  = it.co_no;
        labels.push(it.label);
      });
      picker.innerHTML       = pickerHtml;
      dl.innerHTML           = dlHtml;
      window.PARENT_PROJECT_MAP = newMap;
      var inp = document.getElementById('parent-project-input');
      if (inp && labels.indexOf(inp.value) === -1) inp.value = '';
      picker.value = '';
      syncParentId();
    })
    .catch(function() {});
}

/* keep the hidden parent_id in sync with whatever label is typed/selected —
   this is the actual value saved to the DB, so sub-projects link correctly */
function syncParentId() {
  var inp = document.getElementById('parent-project-input');
  var hid = document.getElementById('parent-project-id-hidden');
  if (!inp || !hid) return;
  var map = window.PARENT_PROJECT_MAP || {};
  var v = inp.value.trim();
  hid.value = map.hasOwnProperty(v) ? map[v] : '';
}

/* ── tab navigation ── */
function switchTab(n) { currentTab = n; updateUI(); }

function updateUI() {
  for (var i = 1; i <= totalTabs; i++) {
    var s = document.getElementById('section-' + i);
    var t = document.getElementById('tab-' + i);
    if (s) s.style.display = (i === currentTab ? 'block' : 'none');
    if (t) t.classList.toggle('tab-active', i === currentTab);
    var pill = document.getElementById('pill-' + i);
    if (pill) pill.className = 'step-pill' + (i < currentTab ? ' done' : (i === currentTab ? ' active' : ''));
  }
  var bar = document.getElementById('global-save-bar');
  if (bar) bar.style.display = 'flex';
}

/* ── field helpers ── */
function setFieldsReadOnly(readonly) {
  var TRACKED = ['order_value','variation_amt','revised_order_value','baseline_eac','baseline_margin',
    'rev_latest_eac','rev_latest_margin','physical_completion','billed_amount','collection_till_date',
    'itd_cost','bill_submission_pending_date','last_billed_date','project_status','business_line',
    'project_manager','client_name','order_number','order_issue_date','co_no','project_name'];
  TRACKED.forEach(function(nm) {
    var el = document.querySelector('[name="' + nm + '"]');
    if (!el) return;
    if (readonly) {
      el.setAttribute('readonly','readonly'); el.setAttribute('tabindex','-1');
      el.style.background = '#f2f4f6'; el.style.color = '#6b7280';
      el.style.cursor = 'not-allowed'; el.style.pointerEvents = 'none';
      el.style.borderColor = '#e5e7eb';
      if (el.tagName === 'SELECT') el.style.opacity = '0.7';
    } else {
      el.removeAttribute('readonly'); el.removeAttribute('tabindex');
      el.style.background = ''; el.style.color = ''; el.style.cursor = '';
      el.style.pointerEvents = ''; el.style.borderColor = ''; el.style.opacity = '';
    }
  });
  var banner = document.getElementById('readonly-banner');
  if (banner) banner.style.display = readonly ? 'flex' : 'none';
  var picker = document.getElementById('parent-picker-wrap');
  if (picker) picker.style.display = readonly ? 'block' : 'none';
}

function setType(type) {
  var rbParent = document.getElementById('rb-parent');
  var rbChild  = document.getElementById('rb-child');
  var lblP     = document.getElementById('lbl-parent');
  var lblC     = document.getElementById('lbl-child');
  var wrap     = document.getElementById('parent-field-wrap');
  if (rbParent) rbParent.checked = (type === 'parent');
  if (rbChild)  rbChild.checked  = (type === 'child');
  var dark = '#1a2540', white = '#ffffff', plain = '#000e24';
  if (lblP) { lblP.style.background = (type==='parent') ? dark : white; lblP.style.color = (type==='parent') ? white : plain; }
  if (lblC) { lblC.style.background = (type==='child')  ? dark : white; lblC.style.color = (type==='child')  ? white : plain; }
  if (wrap) wrap.style.display = (type === 'child') ? 'block' : 'none';
  setFieldsReadOnly(false);
  if (type === 'child') syncParentId();
}

function onParentPickerChange(sel) {
  var parentName = sel.value;
  if (!parentName) return;
  fetchAndFillParent(parentName);
}

var _parentFillCache = {};
function fetchAndFillParent(parentName) {
  if (!parentName) return;
  if (_parentFillCache[parentName]) { fillFields(_parentFillCache[parentName]); return; }
  fetch('/get_parent_data?parent_name=' + encodeURIComponent(parentName) + '&exclude_id=0')
    .then(function(r) { return r.json(); })
    .then(function(d) { if (d && d.parent_id) { _parentFillCache[parentName] = d; fillFields(d); } })
    .catch(function() {});
}

function fillFields(d) {
  function setVal(nm, val) {
    var el = document.querySelector('[name="' + nm + '"]');
    if (!el || val === undefined || val === null) return;
    el.value = val;
  }
  setVal('co_no', d.parent_co_no); setVal('project_name', d.parent_name);
  setVal('business_line', d.parent_business_line); setVal('project_manager', d.parent_project_manager);
  setVal('client_name', d.parent_client_name); setVal('order_number', d.parent_order_number);
  setVal('order_issue_date', d.parent_order_issue_date); setVal('project_status', d.parent_status);
  setVal('physical_completion', d.parent_physical_completion); setVal('baseline_eac', d.parent_baseline_eac);
  setVal('baseline_margin', d.parent_baseline_margin); setVal('rev_latest_margin', d.parent_rev_latest_margin);
  setVal('order_value', d.parent_order_value); setVal('rev_latest_eac', d.parent_eac);
  setVal('billed_amount', d.parent_billed); setVal('collection_till_date', d.parent_collection);
  setVal('itd_cost', d.parent_itd_cost);
  setTimeout(calcVariation, 10);
}

function calcVariation() {
  var ovEl = document.querySelector('[name="order_value"]');
  var vaEl = document.querySelector('[name="variation_amt"]');
  var rov  = document.getElementById('revised_order_value');
  if (!ovEl || !vaEl || !rov) return;
  var ov = parseFloat(ovEl.value) || 0;
  var va = parseFloat(vaEl.value) || 0;
  rov.value = (ov + va).toFixed(2);
  calcMargins();
}
document.addEventListener('input',  function(e) { if (e.target.name==='order_value'||e.target.name==='variation_amt') calcVariation(); });
document.addEventListener('change', function(e) { if (e.target.name==='order_value'||e.target.name==='variation_amt') calcVariation(); });

/* ── Margins: (1 - (Baseline/Rev. Latest EAC ÷ Order Value (Revised, if set))) × 100 ── */
function calcMargins() {
  var ovEl  = document.querySelector('[name="order_value"]');
  var rovEl = document.getElementById('revised_order_value');
  var beEl  = document.getElementById('baseline_eac');
  var rleEl = document.getElementById('rev_latest_eac');
  var bmEl  = document.getElementById('baseline_margin');
  var rlmEl = document.getElementById('rev_latest_margin');
  if (!bmEl || !rlmEl) return;
  var ov  = ovEl  ? (parseFloat(ovEl.value)  || 0) : 0;
  var rov = rovEl ? (parseFloat(rovEl.value) || 0) : 0;
  var ovUsed = rov || ov;
  var be  = beEl  ? (parseFloat(beEl.value)  || 0) : 0;
  var rle = rleEl ? (parseFloat(rleEl.value) || 0) : 0;
  bmEl.value  = ovUsed ? ((1 - (be  / ovUsed)) * 100).toFixed(2) : '0.00';
  rlmEl.value = ovUsed ? ((1 - (rle / ovUsed)) * 100).toFixed(2) : '0.00';
}
document.addEventListener('input',  function(e) { if (e.target.id==='baseline_eac'||e.target.id==='rev_latest_eac') calcMargins(); });
document.addEventListener('change', function(e) { if (e.target.id==='baseline_eac'||e.target.id==='rev_latest_eac') calcMargins(); });

/* ── Extension date helpers ── */
function addMonthsJS(dateStr, months) {
  if (!dateStr) return '';
  try {
    var d = new Date(dateStr);
    if (isNaN(d)) return '';
    var day = d.getDate();
    d.setMonth(d.getMonth() + parseInt(months, 10));
    // If month rolled over (e.g., Jan 31 + 1 month -> Mar 3), fix to last day of target month
    if (d.getDate() !== day) {
      d.setDate(0);
    }
    return d.toISOString().slice(0,10);
  } catch (e) { return ''; }
}

function updateExtendedDisplay() {
  var endEl = document.querySelector('[name="end_date"]');
  var extEl = document.querySelector('[name="extension_months"]');
  var out   = document.getElementById('extended_end_date');
  if (!endEl || !extEl || !out) return;
  // Don't clobber a date the user typed in manually.
  if (out.dataset.manuallyEdited === '1') return;
  var endVal = endEl.value || '';
  var months = parseInt(extEl.value, 10) || 0;
  if (!endVal || !months) { return; }
  out.value = addMonthsJS(endVal, months);
}

/* ── Period of Months -> End Date (auto-computed, hidden) ── */
function updateEndDateFromMonths() {
  var oid = document.querySelector('[name="order_issue_date"]');
  var em  = document.getElementById('end_months');
  var ed  = document.getElementById('end_date');
  if (!oid || !em || !ed) return;
  var months = parseInt(em.value, 10);
  if (!oid.value || isNaN(months)) return;
  ed.value = addMonthsJS(oid.value, months);
  try { updateExtendedDisplay(); } catch (e) {}
}
document.addEventListener('input', function(e) {
  if (e.target && (e.target.id === 'end_months' || e.target.name === 'order_issue_date')) {
    updateEndDateFromMonths();
  }
});

document.addEventListener('input', function(e) {
  if (e.target && e.target.id === 'extended_end_date') {
    // User is typing the extended end date directly -> stop auto-overwriting it.
    e.target.dataset.manuallyEdited = '1';
  }
});

document.addEventListener('input', function(e) { if (e.target && e.target.name === 'extension_months') updateExtendedDisplay(); });

/* ── Validation toast helper ── */
function showValidationToast(msg) {
  var existing = document.getElementById('validation-toast');
  if (existing) existing.remove();
  var toast = document.createElement('div');
  toast.id = 'validation-toast';
  toast.style.cssText = [
    'position:fixed', 'top:68px', 'left:50%', 'transform:translateX(-50%)',
    'z-index:9999', 'background:#fff1f0', 'color:#b91c1c',
    'border:1.5px solid #fca5a5', 'border-radius:10px',
    'padding:12px 22px 12px 16px', 'font-size:13px', 'font-weight:600',
    'box-shadow:0 4px 24px rgba(185,28,28,.18)',
    'display:flex', 'align-items:center', 'gap:10px',
    'max-width:420px', 'pointer-events:none',
    'animation:toast-in .2s ease'
  ].join(';');
  toast.innerHTML = '<span style="font-size:18px">⚠️</span><span>' + msg + '</span>';
  if (!document.getElementById('toast-keyframes')) {
    var style = document.createElement('style');
    style.id = 'toast-keyframes';
    style.textContent = '@keyframes toast-in{from{opacity:0;transform:translateX(-50%) translateY(-10px)}to{opacity:1;transform:translateX(-50%) translateY(0)}}';
    document.head.appendChild(style);
  }
  document.body.appendChild(toast);
  setTimeout(function() {
    toast.style.opacity = '0';
    toast.style.transition = 'opacity .3s';
    setTimeout(function() { if (toast.parentNode) toast.remove(); }, 320);
  }, 3500);
}

function highlightField(name) {
  var el = document.querySelector('[name="' + name + '"]');
  if (!el) return;
  var origBorder = el.style.borderColor;
  var origBg     = el.style.background;
  el.style.borderColor = '#ef4444';
  el.style.background  = '#fff5f5';
  el.focus();
  setTimeout(function() {
    el.style.borderColor = origBorder;
    el.style.background  = origBg;
  }, 3000);
}

/* ── GLOBAL SAVE ── */
function globalSave(goToDocuments) {
  var coEl   = document.querySelector('[name="co_no"]');
  var nameEl = document.querySelector('[name="project_name"]');
  var coVal   = coEl   ? coEl.value.trim()   : '';
  var nameVal = nameEl ? nameEl.value.trim() : '';

  if (!coVal && !nameVal) {
    switchTab(1);
    showValidationToast('Please fill in CO. No. and Project Name before saving.');
    highlightField('co_no'); highlightField('project_name');
    return;
  }
  if (!coVal) {
    switchTab(1);
    showValidationToast('CO. No. is required — please enter a value.');
    highlightField('co_no'); return;
  }
  if (!nameVal) {
    switchTab(1);
    showValidationToast('Project Name is required — please enter a value.');
    highlightField('project_name'); return;
  }

  var btn   = document.getElementById('gsave-btn');
  var label = document.getElementById('gsave-label');
  if (btn) btn.disabled = true;
  if (label) label.innerHTML = '⏳ Saving…';

  var form = document.getElementById('main-project-form');
  var fd   = new FormData(form);

  fetch('/submit', {
    method : 'POST',
    body   : fd,
    headers: { 'X-Requested-With': 'XMLHttpRequest' }
  })
  .then(function(r) {
    if (!r.ok) {
      return r.text().then(function(t) {
        var msg = 'Save failed (' + r.status + ')';
        try { var j = JSON.parse(t); if (j && j.error) msg += ': ' + j.error; } catch (e) {}
        throw new Error(msg);
      });
    }
    return r.json();
  })
  .then(function(d) {
    if (d && d.id) {
      var idInput = form.querySelector('input[name="id"]');
      if (!idInput) {
        idInput = document.createElement('input');
        idInput.type = 'hidden'; idInput.name = 'id';
        form.appendChild(idInput);
      }
      idInput.value = d.id;
      var docPid = document.getElementById('doc-project-id');
      if (docPid) docPid.value = d.id;
    }

    var docName   = (document.querySelector('[name="doc_name"]')   || {}).value || '';
    var docCat    = (document.querySelector('[name="doc_category"]')|| {}).value || '';
    var gdrive    = (document.querySelector('[name="gdrive_link"]') || {}).value || '';
    var netPath   = (document.querySelector('[name="net_path"]')    || {}).value || '';
    var projectId = d ? d.id : (form.querySelector('input[name="id"]') || {}).value || '';

    if ((docName || gdrive || netPath) && projectId) {
      var docFd = new FormData();
      docFd.append('project_id',   projectId);
      docFd.append('doc_name',     docName);
      docFd.append('doc_category', docCat);
      docFd.append('gdrive_link',  gdrive);
      docFd.append('net_path',     netPath);
      return fetch('/link_doc', {
        method : 'POST', body: docFd,
        headers: { 'X-Requested-With': 'XMLHttpRequest' }
      }).then(function() { return d; });
    }
    return d;
  })
  .then(function(d) {
    if (label) label.innerHTML = '✔ Saved!';
    clearDirty();
    setTimeout(function() {
      var pid = d && d.id ? d.id : '';
      var tabNum = goToDocuments ? 4 : currentTab;
      var url = '/data_entry' + (pid ? '?id=' + pid : '') + '#tab' + tabNum;
      window.location.href = url;
    }, 400);
  })
  .catch(function(err) {
    if (btn)  btn.disabled = false;
    if (label) label.innerHTML = '💾 Save';
    alert('Could not save.\\n' + err.message + '\\n\\nCheck required fields and try again.');
  });
}

function saveAll()        { globalSave(false); }
function saveAndGoToDocs(){ globalSave(true);  }

/* ── Toggle collapsible panels ── */
function togglePanel(panelId, btnId) {
  var panel = document.getElementById(panelId);
  var btn   = document.getElementById(btnId);
  if (!panel) return;
  var isHidden = panel.style.display === 'none' || panel.style.display === '';
  panel.style.display = isHidden ? 'block' : 'none';
  if (btn) {
    btn.classList.toggle('open', isHidden);
    /* update label arrow */
    var arrow = btn.querySelector('.panel-arrow');
    if (arrow) arrow.textContent = isHidden ? '▲' : '▼';
  }
}

window.onload = function() {
  var hashMatch = window.location.hash.match(/^#tab([1-4])$/);
  if (hashMatch) currentTab = parseInt(hashMatch[1], 10);
  updateUI();

  var rbChild = document.getElementById('rb-child');
  var initialType = (rbChild && rbChild.checked) ? 'child' : 'parent';
  var isEditing = !!(document.querySelector('input[name="id"]') && document.querySelector('input[name="id"]').value);
  if (isEditing) {
    var wrap = document.getElementById('parent-field-wrap');
    if (wrap) wrap.style.display = (initialType === 'child') ? 'block' : 'none';
    var rb = document.getElementById('rb-' + initialType);
    if (rb) rb.checked = true;
    var lblP = document.getElementById('lbl-parent');
    var lblC = document.getElementById('lbl-child');
    var dark = '#1a2540', white = '#ffffff', plain = '#000e24';
    if (lblP) { lblP.style.background = (initialType==='parent') ? dark : white; lblP.style.color = (initialType==='parent') ? white : plain; }
    if (lblC) { lblC.style.background = (initialType==='child')  ? dark : white; lblC.style.color = (initialType==='child')  ? white : plain; }
  } else {
    setType('parent');
  }
  // If editing an existing project, show the extension field (Period of
  // Months itself stays fully editable — no more view-only lock).
  if (isEditing) {
    var extWrap = document.getElementById('extension-wrap');
    if (extWrap) extWrap.style.display = 'block';
    try { updateExtendedDisplay(); } catch(e) {}
  } else {
    var extWrap = document.getElementById('extension-wrap');
    if (extWrap) extWrap.style.display = 'none';
  }
  setFieldsReadOnly(false);
  try { calcMargins(); } catch(e) {}

  // If we arrived with a ?status= filter (e.g. from the "Active Projects"
  // KPI card on the Dashboard), open the All Projects panel and apply it.
  var statusFilterEl = document.getElementById('proj-status-filter');
  if (statusFilterEl && statusFilterEl.value) {
    var panel = document.getElementById('all-projects-panel');
    var btn   = document.getElementById('btn-all-projects');
    if (panel && panel.style.display === 'none') togglePanel('all-projects-panel', 'btn-all-projects');
    filterProjects();
  }
};
</script>"""

    FOLDER_JS = """<script>
function openServerFolder(path) {
  fetch('/open_folder?path=' + encodeURIComponent(path))
    .then(function(r) { return r.json(); })
    .then(function(d) { if (!d.ok) alert('Could not open folder:\\n' + (d.error || 'Unknown error')); })
    .catch(function(e) { alert('Server error: ' + e); });
}
function openDE_Folder(path) {
  fetch('/open_folder?path=' + encodeURIComponent(path))
    .then(function(r) { return r.json(); })
    .then(function(d) {
      if (!d.ok) {
        navigator.clipboard.writeText(path).then(function() {
          alert('Could not open folder directly.\\nPath copied to clipboard:\\n' + path);
        });
      }
    })
    .catch(function() { navigator.clipboard.writeText(path); });
}
</script>"""

    FILTER_JS = """<script>
function filterProjects() {
  var search = document.getElementById('proj-search').value.toLowerCase().trim();
  var lob    = document.getElementById('proj-lob-filter').value.toLowerCase();
  var status = document.getElementById('proj-status-filter').value.toLowerCase();
  var rows   = document.querySelectorAll('#proj-tbody tr');
  var visible = 0;
  rows.forEach(function(row) {
    var co     = row.dataset.co     || '';
    var name   = row.dataset.name   || '';
    var rowLob = row.dataset.lob    || '';
    var rowSt  = row.dataset.status || '';
    var show   = (!search || co.includes(search) || name.includes(search))
              && (!lob    || rowLob === lob)
              && (!status || rowSt  === status);
    row.style.display = show ? '' : 'none';
    if (show) visible++;
  });
  document.getElementById('proj-count').textContent = visible;
  document.getElementById('proj-no-results').style.display = visible === 0 ? 'block' : 'none';
}
function clearProjFilters() {
  document.getElementById('proj-search').value        = '';
  document.getElementById('proj-lob-filter').value    = '';
  document.getElementById('proj-status-filter').value = '';
  filterProjects();
}
</script>"""

    # ── Data helpers ─────────────────────────────────────────────────────────
    business_lines   = get_visible_lob_list()
    project_managers = get_project_managers()
    lob_list         = get_visible_lob_list()
    project_statuses = get_project_statuses()
    doc_categories   = get_doc_categories()
    parent_project_options  = get_parent_project_options()
    parent_projects         = [o["label"] for o in parent_project_options]
    parent_opts             = "".join(f'<option value="{pp}">{pp}</option>' for pp in parent_projects)
    import json as _json_pp
    # Map both the combined "CO_NO-Name" label AND the plain project name to the parent's co_no, so
    # the hidden parent_id field resolves correctly whether the datalist label or the stored plain
    # name (as shown when editing an existing sub-project) is what's currently in the input.
    # NOTE: the hidden field is named "parent_id" for historical reasons, but the value it holds
    # is the parent's co_no (co_no is unique per project), not the internal database id.
    _pp_map = {}
    for o in parent_project_options:
        _pp_map[o["label"]] = o["co_no"]
        _pp_map[o["name"]]  = o["co_no"]
    parent_project_map_json = _json_pp.dumps(_pp_map)
    child_projects   = get_child_project_names()
    clients          = get_client_list()

    def val(k, d=""): return p.get(k, d)
    edit_extension_html = ""

    def _months_between(start_str, end_str):
        """Whole months between two ISO date strings (used to prefill the
        'Period of Months' field when editing an existing project)."""
        try:
            d1 = datetime.fromisoformat(str(start_str)[:10]).date()
            d2 = datetime.fromisoformat(str(end_str)[:10]).date()
        except Exception:
            return ""
        months = (d2.year - d1.year) * 12 + (d2.month - d1.month)
        if d2.day < d1.day:
            months -= 1
        return max(months, 0)

    end_months_val = ""
    if val("end_date") and val("order_issue_date"):
        end_months_val = _months_between(val("order_issue_date"), val("end_date"))

    def _fmt_months(v):
        """Display an extension_months value without a trailing '.0'."""
        try:
            f = float(v)
        except Exception:
            return ""
        return str(int(f)) if f == int(f) else str(f)

    extension_months_val = _fmt_months(val("extension_months")) if val("extension_months") else ""


    def _ref_manager_block(title: str, category: str, values: list) -> str:
        chips = "".join(f"""
        <span style="display:inline-flex;align-items:center;gap:6px;
                      background:{C['surface_low']};border:1px solid {C['outline']};
                      border-radius:14px;padding:4px 10px 4px 12px;font-size:12px;
                      color:{C['on_surface']};margin:0 6px 6px 0">
          {v}
          <button type="button"
            style="font-size:11px;line-height:1;padding:0;border:none;background:none;cursor:pointer;color:#b91c1c"
            onclick="if(confirm('Remove &quot;{v}&quot; from {title}?')){{
              var f=document.createElement('form');f.method='post';f.action='/delete_ref_option';
              var c=document.createElement('input');c.type='hidden';c.name='category';c.value='{category}';
              var n=document.createElement('input');n.type='hidden';n.name='value';n.value='{v}';
              f.appendChild(c);f.appendChild(n);document.body.appendChild(f);f.submit();
            }}">✕</button>
        </span>""" for v in values) or '<p style="font-size:12px;color:#aaa;margin-bottom:8px">No values yet.</p>'

        return f"""
    <div style="border:1px solid {C['outline']};border-radius:8px;padding:14px;background:#fff">
      <p style="font-size:12px;font-weight:700;color:{C['dark']};margin-bottom:10px">{title}</p>
      <div style="margin-bottom:10px">{chips}</div>
      <form method="post" action="/add_ref_option" style="display:flex;gap:6px">
        <input type="hidden" name="category" value="{category}">
        <input type="text"  name="value" placeholder="Add new…" required
               style="flex:1;height:32px;border:1px solid {C['outline']};border-radius:6px;
                      padding:0 10px;font-size:12px;outline:none">
        <button type="submit"
                style="height:32px;padding:0 12px;background:{C['dark']};color:#fff;border:none;
                       border-radius:6px;font-size:11px;font-weight:700;cursor:pointer">+ Add</button>
      </form>
    </div>"""

    def _fi(lbl, name, ph="", ftype="text", el_id=None):
        v = val(name)
        id_attr = f' id="{el_id}"' if el_id else ""
        return _field_html(lbl, f'<input type="{ftype}" name="{name}" placeholder="{ph}" value="{v}"{id_attr}>')

    def _sl(lbl, name, opts, ph="Select…", extra=""):
        return _field_html(lbl, _sel(name, opts, val(name), ph, extra=extra))

    def _sl_autocomplete(lbl, name, opts, ph="Type or select…"):
        dl_id   = f"dl_{name}"
        dl_opts = "".join(f'<option value="{o}">' for o in opts)
        inp = (f'<input type="text" name="{name}" placeholder="{ph}" '
               f'value="{val(name)}" list="{dl_id}" autocomplete="off">'
               f'<datalist id="{dl_id}">{dl_opts}</datalist>')
        return _field_html(lbl, inp)

    def _money(lbl, name, ph="0.00", el_id=None):
        v = val(name)
        id_attr = f' id="{el_id}"' if el_id else ""
        return _field_html(lbl,
            f'<div class="cur-wrap"><span class="pre">₹</span>'
            f'<input type="number" name="{name}" placeholder="{ph}" value="{v}" step="0.01"{id_attr}>'
            f'</div>')

    def _pct(lbl, name, dv="0"):
        v = val(name, dv)
        return _field_html(lbl,
            f'<div class="suf-wrap">'
            f'<input type="number" name="{name}" min="0" max="100" value="{v}">'
            f'<span class="suf">%</span></div>')

    # Folder path field — used in Tab 3 (Status) to store the project's network folder
    folder_path_field_html = _field_html(
        "Folder Path",
        f'<input type="text" name="folder_path" placeholder="\\\\server\\share\\ProjectDocs\\" '
        f'value="{val("folder_path")}" style="font-family:monospace">'
    )

    if p.get("id"):
        edit_extension_html = ""

    # ── Build doc table rows ──────────────────────────────────────────────────
    doc_rows = ""
    all_docs = fetch_documents()
    for r in all_docs:
        doc_id   = r[0]
        doc_name = r[1]
        doc_cat  = r[2]
        raw      = r[3] or ""

        gdrive, net = "", ""
        if "GDRIVE::" in raw or "NET::" in raw:
            for part in raw.split("|"):
                part = part.strip()
                if part.startswith("GDRIVE::"):
                    gdrive = part[8:].strip()
                elif part.startswith("NET::"):
                    net = part[5:].strip()
        elif raw.startswith("http"):
            gdrive = raw
        elif raw:
            net = raw

        if gdrive:
            drive_cell = (
                f'<a href="{gdrive}" target="_blank" '
                f'style="display:inline-flex;align-items:center;gap:5px;color:{C["dark"]};'
                f'font-weight:600;font-size:13px;text-decoration:none;'
                f'border:1px solid {C["outline"]};border-radius:4px;padding:3px 9px;'
                f'background:{C["surface_low"]}">'
                f'{GDRIVE_SVG}Open Drive</a>'
            )
        else:
            drive_cell = '<span style="color:#ccc;font-size:12px">—</span>'

        if net:
            url_net   = net.replace("\\", "/")
            truncated = net[:45] + ("…" if len(net) > 45 else "")
            net_cell  = (
                f'<span onclick="openServerFolder(\'{url_net}\')" title="{net}" '
                f'style="font-family:monospace;font-size:12px;color:{C["dark"]};'
                f'cursor:pointer;text-decoration:underline;text-decoration-style:dotted;'
                f'text-underline-offset:3px;padding:3px 6px;border-radius:4px;'
                f'display:inline-block;transition:background .15s" '
                f'onmouseover="this.style.background=\'{C["surface_high"]}\'" '
                f'onmouseout="this.style.background=\'transparent\'">'
                f'{truncated}</span>'
            )
        else:
            net_cell = '<span style="color:#ccc;font-size:12px">—</span>'

        doc_rows += f"""
        <tr>
          <td style="font-weight:600;color:{C['dark']}">{doc_name}</td>
          <td>{doc_cat}</td>
          <td>{drive_cell}</td>
          <td>{net_cell}</td>
          <td style="text-align:right">
            <button type="button" class="del-btn"
              onclick="if(confirm('Delete this document?')){{
                var f=document.createElement('form');f.method='post';f.action='/delete_doc';
                var i=document.createElement('input');i.type='hidden';i.name='doc_id';i.value='{doc_id}';
                f.appendChild(i);document.body.appendChild(f);f.submit();
              }}">🗑</button>
          </td>
        </tr>"""

    if not doc_rows:
        doc_rows = '<tr><td colspan="5" style="text-align:center;color:#aaa;padding:20px">No documents linked yet.</td></tr>'

    # ── All-projects table rows ───────────────────────────────────────────────
    all_projects    = fetch_all_projects()

    # LOB values actually present on existing/imported projects only (for the
    # "All Projects" filter dropdown) — not the full master ref_lob list.
    _used_lob_set = set()
    for _pr in all_projects:
        _lob_raw = (_pr.get('hier_business_line') or _pr.get('business_line') or '').strip()
        if _lob_raw:
            _used_lob_set.add(_lob_raw)
    used_lob_list = sorted(_used_lob_set, key=lambda s: s.lower())

    proj_rows_html  = ""
    for pr in all_projects:
        lob_val    = (pr.get('hier_business_line') or pr.get('business_line') or '').lower()
        status_val = (pr.get('project_status') or '').lower()
        co_val     = (pr.get('co_no') or '').lower()
        name_val   = (pr.get('project_name') or '').lower()
        pid        = pr.get('id', '')
        proj_docs  = fetch_documents(pid)

        drive_links = []
        for doc in proj_docs:
            try:
                dname = doc.get('doc_name','Drive') if isinstance(doc, dict) else doc[1]
                raw   = (doc.get('doc_link','') if isinstance(doc, dict) else doc[3]) or ''
                if 'GDRIVE::' in raw:
                    for part in raw.split('|'):
                        part = part.strip()
                        if part.startswith('GDRIVE::'):
                            url = part[8:].strip()
                            if url: drive_links.append((dname, url))
                elif raw.startswith('http'):
                    drive_links.append((dname, raw))
            except Exception:
                continue

        if drive_links:
            links_html = " ".join(
                f'<a href="{url}" target="_blank" title="{name}" '
                f'style="display:inline-flex;align-items:center;gap:4px;font-size:11px;font-weight:600;'
                f'color:{C["dark"]};text-decoration:none;border:1px solid {C["outline"]};'
                f'border-radius:4px;padding:2px 8px;background:{C["surface_low"]};'
                f'margin-right:4px;white-space:nowrap">'
                f'{GDRIVE_SVG}{name[:18]}</a>'
                for name, url in drive_links[:2]
            )
            drive_cell = f'<td style="padding:10px 14px">{links_html}</td>'
        else:
            drive_cell = f'<td style="padding:10px 14px;color:#ccc;font-size:11px">—</td>'

        fp = (pr.get('folder_path') or '').strip()
        if not fp:
            for doc in proj_docs:
                try:
                    raw = (doc[3] if not isinstance(doc, dict) else doc.get('doc_link','')) or ''
                    if 'NET::' in raw:
                        for part in raw.split('|'):
                            part = part.strip()
                            if part.startswith('NET::'):
                                fp = part[5:].strip(); break
                    if fp: break
                except Exception:
                    continue

        if fp:
            fp_esc     = fp.replace("\\", "\\\\").replace("'", "\\'")
            fp_display = fp[:35] + '…' if len(fp) > 35 else fp
            folder_cell = (
                f'<td style="padding:10px 14px">'
                f'<span onclick="openDE_Folder(\'{fp_esc}\')" title="{fp}" '
                f'style="font-family:monospace;font-size:11px;color:{C["dark"]};'
                f'cursor:pointer;text-decoration:underline;text-decoration-style:dotted;'
                f'text-underline-offset:3px;padding:2px 5px;border-radius:4px;'
                f'display:inline-block;transition:background .15s" '
                f'onmouseover="this.style.background=\'{C["surface_high"]}\'" '
                f'onmouseout="this.style.background=\'transparent\'">'
                f'📂 {fp_display}</span></td>'
            )
        else:
            folder_cell = f'<td style="padding:10px 14px;color:#ccc;font-size:11px">—</td>'

        proj_rows_html += f"""
        <tr data-lob="{lob_val}" data-status="{status_val}"
            data-co="{co_val}" data-name="{name_val}">
          <td style="padding:10px 14px;font-size:12px;color:{C['on_var']}">{pr.get('co_no','')}</td>
          <td style="padding:10px 14px;font-weight:600;color:{C['dark']};font-size:13px">{pr.get('project_name','')}</td>
          <td style="padding:10px 14px;font-size:12px;color:{C['on_var']}">{pr.get('project_manager','')}</td>
          <td style="padding:10px 14px;font-size:12px">{pr.get('project_status','')}</td>
          {drive_cell}
          {folder_cell}
          <td style="padding:10px 14px;text-align:right">
            <a href="/edit_project?id={pr['id']}"
               style="color:{C['dark']};font-size:12px;font-weight:700;text-decoration:none">✏ Edit</a>
            &nbsp;
            <button class="del-btn" type="button"
              onclick="if(confirm('Delete this project?')){{
                var f=document.createElement('form');f.method='post';f.action='/delete_project';
                var i=document.createElement('input');i.type='hidden';i.name='project_id';i.value='{pr['id']}';
                f.appendChild(i);document.body.appendChild(f);f.submit();
              }}">🗑</button>
          </td>
        </tr>"""

    if not proj_rows_html:
        proj_rows_html = '<tr><td colspan="7" style="text-align:center;color:#aaa;padding:18px">No projects yet.</td></tr>'

    project_id_hidden = f'<input type="hidden" name="id" value="{val("id")}">' if val("id") else ""

    # ══════════════════════════════════════════════════════════════════════════
    # BUILD HTML
    # ══════════════════════════════════════════════════════════════════════════
    html  = _html_head("Project Data Entry") + ENTRY_CSS + "<body>"
    html += _sidebar("data_entry") + _topbar("PROJECT DATA ENTRY", username)

    html += f"""
<div class="main"><div class="canvas">
  <h2 class="page-title">Project Data Entry</h2>
  <p class="page-sub">Log professional surveyor project data with precision and accuracy.</p>
  {alert_html}

  <!-- ── UPLOAD ZONE ── -->
  <form method="post" action="/upload" enctype="multipart/form-data">
    <div class="upload-zone" onclick="document.getElementById('fi').click()">
      <div style="width:52px;height:52px;background:{C['dark']};border-radius:50%;display:flex;
                  align-items:center;justify-content:center;margin:0 auto 12px;font-size:24px;color:#fff">📤</div>
      <h3 style="font-size:17px;font-weight:700;color:{C['dark']};margin-bottom:6px">Excel / CSV Upload</h3>
      <p style="font-size:13px;color:#6b7280;margin-bottom:12px">Drag and drop or
         <strong style="color:{C['dark']};text-decoration:underline">browse files</strong></p>
      <div style="display:inline-flex;gap:12px;background:{C['surface_high']};padding:5px 14px;
                  border-radius:20px;font-size:11px;font-weight:700;letter-spacing:.07em;color:{C['on_var']}">
        XLSX &nbsp;/&nbsp; CSV &nbsp;/&nbsp; MAX 25MB
      </div>
      <input id="fi" type="file" name="project_file" accept=".xlsx,.csv"
             style="display:none" onchange="this.form.submit()">
    </div>
  </form>

  <!-- ── Export / action bar ── -->
  <div class="exp-bar" style="display:flex;justify-content:space-between;align-items:center">
    <div style="display:flex;gap:10px;flex-wrap:wrap;align-items:center">
      <a class="exp-btn" href="/export_excel">⬇ Export Excel</a>

      <!-- Clear memory / data button -->
      <form method="post" action="/wipe_data" style="display:inline;margin:0"
            onsubmit="return confirm('Are you sure you want to completely clear all project, expenditure, and LOB data from database memory? This will reset the database so you can upload a fresh Excel file.');">
        <button type="submit" class="exp-btn" style="color:#ef4444;border-color:#fca5a5;background:#fff" title="Wipe previous data and LOBs from memory">
          🗑 Clear Memory / Data
        </button>
      </form>

      <!-- All Projects toggle -->
      <button type="button" id="btn-all-projects" class="exp-btn toggle-panel-btn"
              onclick="togglePanel('all-projects-panel','btn-all-projects')"
              style="display:inline-flex;align-items:center;gap:6px">
        📋 All Projects ({len(all_projects)}) <span class="panel-arrow">▼</span>
      </button>

      <!-- Manage Dropdowns toggle -->
      <button type="button" id="btn-manage-dropdowns" class="exp-btn toggle-panel-btn"
              onclick="togglePanel('manage-dropdowns-panel','btn-manage-dropdowns')"
              style="display:inline-flex;align-items:center;gap:6px">
        ⚙ Manage Dropdowns <span class="panel-arrow">▼</span>
      </button>
    </div>

    <div style="display:flex;gap:10px;align-items:center">
      <a href="/data_entry"
         style="background:{C['dark']};color:#fff;border:none;padding:10px 28px;
                border-radius:8px;font-size:13px;font-weight:700;text-transform:uppercase;
                letter-spacing:.05em;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.20);
                display:inline-flex;align-items:center;gap:8px;transition:opacity .15s;
                text-decoration:none">
        ＋ New Project
      </a>
      <button type="button" id="top-save-btn"
              onclick="saveAll()"
              style="background:{C['dark']};color:#fff;border:none;padding:10px 28px;
                     border-radius:8px;font-size:13px;font-weight:700;text-transform:uppercase;
                     letter-spacing:.05em;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.20);
                     display:inline-flex;align-items:center;gap:8px;transition:opacity .15s">
        💾 Save All
      </button>
    </div>
  </div>

  <!-- ══ ALL PROJECTS COLLAPSIBLE PANEL ══ -->
  <div id="all-projects-panel" style="display:none;margin-bottom:20px">
    <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;overflow:hidden">
      <div style="padding:12px 18px;border-bottom:1px solid {C['outline']};
                  background:{C['surface_low']};display:flex;align-items:center;
                  justify-content:space-between;flex-wrap:wrap;gap:10px">
        <span style="font-size:13px;font-weight:700;color:{C['dark']}">
          All Projects (<span id="proj-count">{len(all_projects)}</span>)
        </span>
        <a href="/data_entry"
           style="background:{C['dark']};color:#fff;border-radius:6px;padding:5px 12px;
                  font-size:11px;text-decoration:none;font-weight:700">+ New Project</a>
      </div>
      <!-- Filter bar -->
      <div style="padding:10px 18px;border-bottom:1px solid {C['outline']};background:#fff;
                  display:flex;gap:10px;flex-wrap:wrap;align-items:center">
        <div style="position:relative;flex:2;min-width:180px">
          <span style="position:absolute;left:10px;top:50%;transform:translateY(-50%);
                       font-size:13px;color:#9ca3af;pointer-events:none">🔍</span>
          <input id="proj-search" type="text" placeholder="Search by project name or CO.No…"
                 oninput="filterProjects()"
                 style="width:100%;padding:7px 10px 7px 30px;border:1px solid {C['outline']};
                        border-radius:6px;font-size:12px;outline:none;box-sizing:border-box">
        </div>
        <select id="proj-lob-filter" onchange="filterProjects()"
                style="flex:1;min-width:140px;padding:7px 10px;border:1px solid {C['outline']};
                       border-radius:6px;font-size:12px;background:#fff;outline:none;color:{C['dark']};cursor:pointer">
          <option value="">All LOB</option>
          {"".join(f'<option value="{lo}">{lo}</option>' for lo in used_lob_list)}
        </select>
        <select id="proj-status-filter" onchange="filterProjects()"
                style="flex:1;min-width:120px;padding:7px 10px;border:1px solid {C['outline']};
                       border-radius:6px;font-size:12px;background:#fff;outline:none;color:{C['dark']};cursor:pointer">
          <option value="" {"selected" if not status_filter else ""}>All Status</option>
          {"".join(f'<option value="{s}" {"selected" if s == status_filter else ""}>{s}</option>' for s in project_statuses)}
        </select>
        <button type="button" onclick="clearProjFilters()"
                style="padding:7px 12px;border:1px solid {C['outline']};border-radius:6px;
                       background:#fff;font-size:11px;font-weight:600;cursor:pointer;
                       color:{C['on_var']};white-space:nowrap">✕ Clear</button>
      </div>
      <div style="overflow-x:auto">
        <table class="proj-table" id="proj-table">
          <thead><tr>
            <th>CO. No.</th><th>Project Name</th><th>Manager</th><th>Status</th>
            <th>Google Drive</th><th>Folder Path</th>
            <th style="text-align:right">Actions</th>
          </tr></thead>
          <tbody id="proj-tbody">{proj_rows_html}</tbody>
        </table>
      </div>
      <div id="proj-no-results"
           style="display:none;padding:24px;text-align:center;font-size:13px;color:#aaa">
        No projects match your filters.
      </div>
    </div>
  </div><!-- /all-projects-panel -->

  <!-- ══ MANAGE DROPDOWNS COLLAPSIBLE PANEL ══ -->
  <div id="manage-dropdowns-panel" style="display:none;margin-bottom:20px">
    <div class="card" style="margin-bottom:0">
      <div style="padding:14px 20px;border-bottom:1px solid {C['outline']};background:{C['surface_low']}">
        <span style="font-size:13px;font-weight:700;color:{C['dark']}">⚙ Manage Dropdown Options</span>
        <span style="font-size:12px;color:{C['on_var']};margin-left:8px">Add / remove values from all dropdowns</span>
      </div>
      <div style="padding:20px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px">
        {_ref_manager_block("Line of Business (LOB)",  "lob",    lob_list)}
        {_ref_manager_block("Business Lines",          "biz",    business_lines)}
        {_ref_manager_block("Project Managers",        "pm",     project_managers)}
        {_ref_manager_block("Project Statuses",        "status", project_statuses)}
        {_ref_manager_block("Document Categories",     "doccat", doc_categories)}
        {_ref_manager_block("Expense Categories",      "expcat", get_expense_categories())}
      </div>
    </div>
  </div><!-- /manage-dropdowns-panel -->

  <!-- ══════════════════════════════════════════════════════════════════════
       ONE MASTER FORM — wraps ALL four tabs including Documents
       ══════════════════════════════════════════════════════════════════════ -->
  <div class="card">

    <div class="tab-bar">
      <button class="tab-btn" id="tab-1" onclick="switchTab(1)">1. Project Details</button>
      <button class="tab-btn" id="tab-2" onclick="switchTab(2)">2. Cost Plan</button>
      <button class="tab-btn" id="tab-3" onclick="switchTab(3)">3. Status</button>
      <button class="tab-btn" id="tab-4" onclick="switchTab(4)">4. Documents</button>
    </div>

    <form method="post" action="/submit" id="main-project-form">
      {project_id_hidden}

      <!-- ════ TAB 1: PROJECT DETAILS ════ -->
      <div class="section" id="section-1">

        <div id="readonly-banner"
             style="display:none;align-items:center;gap:10px;background:#fffbf0;
                    border:1px solid #fed488;border-radius:8px;padding:10px 16px;
                    margin-bottom:16px;font-size:13px;color:#775a19;font-weight:600">
          <span style="font-size:18px">🔒</span>
          <span>Viewing parent project — all fields are read-only.</span>
        </div>

        <!-- Project Type toggle -->
        <div style="margin-top:22px;padding-top:20px;border-top:1px solid {C['outline']}">
          <label style="font-size:11px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;
                        color:{C['on_var']};display:block;margin-bottom:8px">PROJECT TYPE</label>
          <div style="display:flex;gap:0;border:1px solid {C['outline']};border-radius:8px;overflow:hidden;width:fit-content">
            <label id="lbl-parent"
                   style="display:flex;align-items:center;gap:8px;padding:10px 24px;cursor:pointer;
                          font-size:13px;font-weight:700;transition:all .2s"
                   onclick="setType('parent')">
              <input type="radio" name="project_type" value="parent" id="rb-parent" style="width:15px;height:15px">
              Parent Project
            </label>
            <label id="lbl-child"
                   style="display:flex;align-items:center;gap:8px;padding:10px 24px;cursor:pointer;
                          font-size:13px;font-weight:700;transition:all .2s"
                   onclick="setType('child')">
              <input type="radio" name="project_type" value="child" id="rb-child" style="width:15px;height:15px">
              Sub Project
            </label>
          </div>
        </div>

        <!-- Hierarchy & Classification -->
        <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;padding:22px;margin:20px 0">
          <div class="sec-title" style="margin-bottom:16px">Project Hierarchy &amp; Classification</div>
          <div id="parent-picker-wrap" style="display:none;margin-bottom:16px">
            <div style="background:#f0f4ff;border:1px solid #c8d8f8;border-radius:8px;padding:14px 16px">
              <p style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;
                        color:#1e3a8a;margin-bottom:8px">SELECT PARENT PROJECT TO VIEW</p>
              <select id="parent-picker-sel" onchange="onParentPickerChange(this)"
                      style="min-width:260px;padding:9px 12px;border:1.5px solid #93c5fd;
                             border-radius:6px;font-size:13px;font-weight:600;color:#1e3a8a;
                             background:#fff;outline:none;cursor:pointer">
                <option value="">— Select a parent project —</option>
                {parent_opts}
              </select>
            </div>
          </div>
          <div class="g3">
            {_sl("Line of Business (Hierarchy)","hier_business_line", lob_list, "Select LOB…", extra='id="sel-hier-lob" onchange="filterParentsByLob(this.value)"')}
            <div id="parent-field-wrap">
              <div class="field">
                <label style="font-size:11px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:{C['on_var']}">Parent Project</label>
                <input type="text" name="parent_project" id="parent-project-input"
                       placeholder="Select or type parent…"
                       value="{val('parent_project')}" list="dl_parent_project" autocomplete="off"
                       oninput="syncParentId()" onchange="syncParentId()">
                <input type="hidden" name="parent_id" id="parent-project-id-hidden" value="{val('parent_id')}">
                <datalist id="dl_parent_project">
                  {''.join(f'<option value="{pp}">' for pp in parent_projects)}
                </datalist>
              </div>
            </div>
          </div>
        </div>
        <script>
          window.PARENT_PROJECT_MAP = {parent_project_map_json};
        </script>

        <!-- Core Identification -->
        <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;padding:22px;margin-bottom:20px">
          <div class="sec-title" style="margin-bottom:16px">Core Identification</div>
          <div class="g12">
            {_fi("CO. NO.","co_no","e.g. CO-8872")}
            {_fi("Project Name","project_name","e.g. Western Railway Bridge Inspection")}
          </div>
          <div class="g3 mt">
            {_sl("Line of Business","business_line", business_lines, "Select Business Line…")}
            {_sl("Project Manager","project_manager", project_managers, "Select Manager…")}
            {_sl_autocomplete("Client Name","client_name", clients, "e.g. National Highways Authority")}
          </div>
          <div class="g3 mt">
            {_fi("Order Number","order_number","e.g. ORD-8872")}
            {_fi("Order Issue Date","order_issue_date","",ftype="date")}
            {_field_html("Period of Months", f'''
              <div class="suf-wrap">
                <input type="number" name="end_months" id="end_months" min="0" step="1"
                       placeholder="e.g. 18" value="{end_months_val}">
                <span class="suf">mo</span>
              </div>
              <input type="hidden" name="end_date" id="end_date" value="{val('end_date')}">
            ''')}
            <div id="extension-wrap" style="display:none">
              {_field_html("Extension Months", f'''
                <div class="suf-wrap">
                  <input type="number" name="extension_months" id="extension_months" min="0" step="1"
                         placeholder="e.g. 3" value="{extension_months_val}">
                  <span class="suf">mo</span>
                </div>
                <input type="hidden" name="extended_end_date" id="extended_end_date" value="{val('extended_end_date')}">
              ''')}
            </div>
            {_money("Order Value (₹)","order_value")}
          </div>
          <div class="g2 mt">
            <div class="fin-card">
              <p style="font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;
                        color:{C['secondary']};margin-bottom:14px">VARIATION</p>
              {_money("Variation Amount (₹)","variation_amt","0.00")}
            </div>
            <div class="fin-card">
              <p style="font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;
                        color:{C['dark']};margin-bottom:14px">REVISED ORDER VALUE</p>
              <div class="field">
                <label>Revised Order Value (₹)</label>
                <div class="cur-wrap">
                  <span class="pre">₹</span>
                  <input type="number" name="revised_order_value" id="revised_order_value"
                         value="{val('revised_order_value')}" step="0.01" readonly
                         style="padding-left:24px;background:#e8f0fe;color:#1e3a8a;
                                font-weight:700;cursor:not-allowed;border-color:#93c5fd">
                </div>
                <p style="font-size:11px;color:#3b5bdb;margin-top:8px;font-weight:600">
                  = Order Value + Variation Amount
                </p>
              </div>
            </div>
          </div>
        </div>

      </div><!-- /section-1 -->

      <!-- ════ TAB 2: COST PLAN ════ -->
      <div class="section" id="section-2">
        <div class="sec-title">Financial Baseline &amp; Revision</div>
        <div class="g2">
          <div class="fin-card">
            <p style="font-size:11px;font-weight:700;letter-spacing:.06em;color:{C['secondary']};
                      text-transform:uppercase;margin-bottom:14px">BASELINE DATA</p>
            {_money("Baseline EAC (₹)","baseline_eac","0.00","baseline_eac")}
            <div class="mt">
              <div class="field">
                <label>Baseline Margin (%)</label>
                <div class="suf-wrap">
                  <input type="number" name="baseline_margin" id="baseline_margin"
                         value="{val('baseline_margin')}" readonly disabled tabindex="-1"
                         style="background:#e8f0fe;color:#1e3a8a;font-weight:700;
                                cursor:not-allowed;border-color:#93c5fd;pointer-events:none;
                                opacity:1;-webkit-text-fill-color:#1e3a8a">
                  <span class="suf">%</span>
                </div>
                <p style="font-size:11px;color:#3b5bdb;margin-top:8px;font-weight:600">
                  = (1 − (Baseline EAC ÷ Order Value (Revised, if set))) × 100
                </p>
              </div>
            </div>
          </div>
          <div class="fin-card">
            <p style="font-size:11px;font-weight:700;letter-spacing:.06em;color:{C['dark']};
                      text-transform:uppercase;margin-bottom:14px">LATEST REVISION</p>
            {_money("Rev. Latest EAC (₹)","rev_latest_eac","Enter updated EAC","rev_latest_eac")}
            <div class="mt">
              <div class="field">
                <label>Rev. Latest Margin (%)</label>
                <div class="suf-wrap">
                  <input type="number" name="rev_latest_margin" id="rev_latest_margin"
                         value="{val('rev_latest_margin')}" readonly disabled tabindex="-1"
                         style="background:#e8f0fe;color:#1e3a8a;font-weight:700;
                                cursor:not-allowed;border-color:#93c5fd;pointer-events:none;
                                opacity:1;-webkit-text-fill-color:#1e3a8a">
                  <span class="suf">%</span>
                </div>
                <p style="font-size:11px;color:#3b5bdb;margin-top:8px;font-weight:600">
                  = (1 − (Rev. Latest EAC ÷ Order Value (Revised, if set))) × 100
                </p>
              </div>
            </div>
          </div>
        </div>
      </div><!-- /section-2 -->


      <!-- ════ TAB 3: STATUS ════ -->
      <div class="section" id="section-3">
        <div class="sec-title">Current Execution Status</div>
        <div class="g3">
          {_sl("Project Status","project_status", project_statuses, "Select Status…")}
          {_pct("Physical % Completion","physical_completion")}
          {_money("Up to date Billed Amount (₹)","billed_amount")}
        </div>
        <div class="g12 mt">
          {_money("Collection till Date (₹)","collection_till_date")}
          <div></div>
        </div>
         <div class="mt">
          {folder_path_field_html}
          <p style="font-size:11px;color:#6b7280;margin-top:6px">
            Enter the network path to this project's folder.
          </p>
        </div>
        <div class="info-ban">
          <span style="font-size:18px">ℹ️</span>
          <span>Physical completion should align with the Billed Amount to avoid revenue leakage.</span>
        </div>
      </div><!-- /section-3 -->

      <!-- ════ TAB 4: DOCUMENTS ════ -->
      <div class="section" id="section-4">
        <input type="hidden" name="doc_project_id" id="doc-project-id" value="{val('id')}">

        <div class="sec-title">Link New Document</div>
        <div style="background:#f8faff;border:1px solid {C['outline']};border-radius:10px;padding:22px;margin-bottom:24px">
          <div class="g2" style="margin-bottom:16px">
            {_sl("Document Category","doc_category", doc_categories, "Select Category…")}
            {_fi("Document Name","doc_name","e.g. Bridge Inspection Report v1.0")}
          </div>
          <div style="background:#fff;border:1px solid {C['outline']};border-radius:8px;padding:18px">
            <p style="font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;
                      color:{C['on_var']};margin-bottom:14px">📎 Document Location — fill one or both</p>
            <div class="g2">
              <div class="field">
                <label style="display:inline-flex;align-items:center;gap:6px">
                  {GDRIVE_SVG} Google Drive Link
                </label>
                <input type="text" name="gdrive_link"
                       placeholder="https://drive.google.com/…"
                       style="height:40px;border:1px solid {C['outline']};border-radius:6px;
                              padding:0 12px;font-size:14px;width:100%;box-sizing:border-box;outline:none">
              </div>
              <div class="field">
                <label>🗂 Network / Folder Path</label>
                <input type="text" name="net_path"
                       placeholder="\\\\server\\share\\ProjectDocs\\"
                       style="height:40px;border:1px solid {C['outline']};border-radius:6px;
                              padding:0 12px;font-size:14px;font-family:monospace;
                              width:100%;box-sizing:border-box;outline:none">
              </div>
            </div>
          </div>
          <p style="font-size:12px;color:#6b7280;margin-top:10px">
            ℹ️ Fill in document details above and click <strong>Save All</strong> — the document will be linked when the project is saved.
          </p>
        </div>

        <!-- Linked documents table -->
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px">
          <p class="sec-title" style="margin-bottom:0">Linked Documents ({len(all_docs)})</p>
        </div>
        <div style="border:1px solid {C['outline']};border-radius:8px;overflow:hidden">
          <table class="proj-table">
            <thead><tr>
              <th>Document Name</th><th>Category</th>
              <th>Google Drive</th><th>Network Path</th>
              <th style="text-align:right">Action</th>
            </tr></thead>
            <tbody>{doc_rows}</tbody>
          </table>
        </div>
      </div><!-- /section-4 -->

    </form><!-- /main-project-form -->

    <!-- ── GLOBAL STICKY SAVE BAR ── -->
    <div class="global-save-bar" id="global-save-bar">
      <div class="gsb-pills">
        <span class="gsb-dirty" id="gsb-dirty">● Unsaved changes</span>
        <button type="button" class="step-pill" id="pill-1" onclick="switchTab(1)">
          <span class="step-num">1</span> Details
        </button>
        <button type="button" class="step-pill" id="pill-2" onclick="switchTab(2)">
          <span class="step-num">2</span> Cost Plan
        </button>
        <button type="button" class="step-pill" id="pill-3" onclick="switchTab(3)">
          <span class="step-num">3</span> Status
        </button>
        <button type="button" class="step-pill" id="pill-4" onclick="switchTab(4)">
          <span class="step-num">4</span> Documents
        </button>
      </div>
      <button type="button" id="gsave-btn" class="btn-gsave" onclick="saveAll()">
        <span id="gsave-label">💾 Save Project</span>
      </button>
    </div>

  </div><!-- /card -->

</div></div>"""

    html += TAB_JS
    html += FOLDER_JS
    html += FILTER_JS
    return html

# ═══════════════════════════════════════════════════════════════
#  PAGE 3 — EXPENDITURES
# ═══════════════════════════════════════════════════════════════
def _process_expenditure_upload(tmp_path: str) -> str:
    import openpyxl, math, re
    from datetime import datetime

    try:
        wb = openpyxl.load_workbook(tmp_path, read_only=True, data_only=True)
    except Exception as e:
        return f'<div class="alert err">❌ Could not open file: {e}</div>'

    # Find the right sheet (look for "expenditures", "expenditure", "cost", "costs", etc.)
    exp_ws = None
    for sname in wb.sheetnames:
        s_clean = re.sub(r"[^a-z0-9]+", "", sname.lower())
        if any(k in s_clean for k in ("expenditure", "expenditures", "cost", "costs", "expense", "expenses")):
            exp_ws = wb[sname]
            break
    ws = exp_ws or wb.active

    try:
        rows = list(ws.iter_rows(values_only=True))
    except Exception as e:
        return f'<div class="alert err">❌ Could not read sheet rows: {e}</div>'

    if not rows:
        return '<div class="alert err">❌ File is empty.</div>'

    def _norm_header(value):
        if value is None:
            return ""
        s = str(value).strip().lower()
        s = s.replace("&", "and")
        s = re.sub(r"[^a-z0-9]+", "_", s).strip("_")
        return s

    # If the workbook contains a Projects sheet with ITD Cost / Column U,
    # extract those authoritative figures so no project's ITD cost is missed or reduced.
    projects_sheet_itd = {}
    for sname in wb.sheetnames:
        if re.sub(r"[^a-z0-9]+", "", sname.lower()) == "projects":
            try:
                p_ws = wb[sname]
                p_rows = list(p_ws.iter_rows(values_only=True))
                if p_rows:
                    p_hdr = [_norm_header(c) for c in p_rows[0]]
                    p_co_idx = None
                    p_itd_idx = None
                    for i_h, h in enumerate(p_hdr):
                        if h in ("co_no", "cono", "co", "project_co_no", "project_code"):
                            p_co_idx = i_h
                        if h in ("itd_cost", "total_itd", "total_expenditure", "itd_expenditure", "total_itd_cost", "total_expenses", "itd"):
                            p_itd_idx = i_h
                    if p_co_idx is not None and p_itd_idx is None and len(p_hdr) >= 21:
                        p_itd_idx = 20  # Column U
                    if p_co_idx is not None and p_itd_idx is not None:
                        for prow in p_rows[1:]:
                            if prow and p_co_idx < len(prow) and p_itd_idx < len(prow):
                                co_val = str(prow[p_co_idx] or "").strip()
                                try:
                                    v_raw = prow[p_itd_idx]
                                    itd_val = float(str(v_raw).replace(",", "").replace("₹", "").strip() or 0.0) if v_raw is not None else 0.0
                                except Exception:
                                    itd_val = 0.0
                                if co_val and itd_val > 0:
                                    projects_sheet_itd[co_val.lower()] = itd_val
            except Exception:
                pass
            break

    # Find the header row (look within first 20 rows for key column names)
    hdr_idx = 0
    for i, row in enumerate(rows[:20]):
        if not row:
            continue
        norm_set = {_norm_header(c) for c in row if c is not None}
        if any(k in norm_set for k in (
            "co_no", "project_id", "project_name", "id", "cono", "co",
            "itd_cogs", "cogs", "lob", "project_manager", "pm", "project_title",
            "total_itd", "total_expenditure", "total_cost"
        )):
            hdr_idx = i
            break

    headers = [_norm_header(c) for c in rows[hdr_idx]]

    def _col(*names):
        wanted = {_norm_header(n) for n in names}
        for idx, header in enumerate(headers):
            if header in wanted:
                return idx
        return None

    cm = {
        "project_id":      _col("project_id", "Project ID", "Project Id", "PID", "ID", "proj_id", "proj id"),
        "co_no":           _col("co_no", "CO.No", "CO No", "CO_NO", "CO NO.", "CO", "CONO", "CO Number",
                                "Project CO No", "Project CO", "Job No", "Job", "Job Number", "Project No",
                                "Project Number", "Project Code", "Co. No."),
        "project_name":    _col("project_name", "Project Name", "Project", "Name", "Project Title", "Title", "Project Description"),
        "project_manager": _col("project_manager", "Project Manager", "PM", "Manager", "PM Name", "Project_Manager"),
        "lob":             _col("lob", "LOB", "Business Line", "Line of Business", "Hier Business Line", "Hier_Business_Line", "Business_Line"),
        "data_as_of":      _col("data_as_of", "Data As Of", "As Of Date", "As On Date", "Date", "Data_As_Of", "As Of"),
        "erp_update_date": _col("erp_update_date", "ERP Update Date", "Update Date", "ERP Date", "ERP_Update_Date"),
        "itd_cogs":        _col("itd_cogs", "ITD COGS", "ITD_COGS", "COGS ITD", "COGS (ITD)", "COGS", "ITD Cost COGS"),
        "itd_djcs":        _col("itd_djcs", "ITD DJCS", "ITD_DJCS", "DJCS ITD", "DJCS (ITD)", "DJCS"),
        "itd_equip":       _col("itd_equip", "ITD Equipment", "ITD Equip", "ITD_Equip", "ITD_Equipment", "Equipment ITD", "Equipment (ITD)", "Equipment", "Equip"),
        "itd_manpower":    _col("itd_manpower", "ITD Manpower", "ITD_Manpower", "Manpower ITD", "Manpower (ITD)", "Manpower"),
        "itd_vehicles":    _col("itd_vehicles", "ITD Vehicles", "ITD Vehicle", "ITD_Vehicles", "Vehicles ITD", "Vehicles (ITD)", "Vehicles", "Vehicle"),
        "itd_others":      _col("itd_others", "ITD Others", "ITD Other", "ITD_Others", "Others ITD", "Others (ITD)", "Others", "Other", "Misc", "Miscellaneous"),
        "total_itd":       _col("total_itd", "Total ITD", "Total_ITD", "ITD Total", "Total Expenditure", "Total Cost", "ITD Cost", "Total ITD Cost", "Total_Expenditure", "Total Expenses", "ITD Expenditure", "ITD"),
        "erp_cogs":        _col("erp_cogs", "ERP COGS", "ERP_COGS", "COGS ERP", "COGS (ERP)"),
        "erp_djcs":        _col("erp_djcs", "ERP DJCS", "ERP_DJCS", "DJCS ERP", "DJCS (ERP)"),
        "erp_equip":       _col("erp_equip", "ERP Equipment", "ERP Equip", "ERP_Equip", "ERP_Equipment", "Equipment ERP", "Equipment (ERP)"),
        "erp_manpower":    _col("erp_manpower", "ERP Manpower", "ERP_Manpower", "Manpower ERP", "Manpower (ERP)"),
        "erp_vehicles":    _col("erp_vehicles", "ERP Vehicles", "ERP Vehicle", "ERP_Vehicles", "Vehicles ERP", "Vehicles (ERP)"),
        "erp_others":      _col("erp_others", "ERP Others", "ERP Other", "ERP_Others", "Others ERP", "Others (ERP)"),
        "total_erp":       _col("total_erp", "Total ERP", "Total_ERP", "ERP Total", "Total ERP Cost"),
    }
    if cm.get("total_itd") is None and len(headers) >= 21:
        if any(term in headers[20] for term in ("itd", "cost", "exp")):
            cm["total_itd"] = 20

    def _num(row, key):
        idx = cm.get(key)
        if idx is None or idx >= len(row):
            return 0.0
        v = row[idx]
        if v is None or isinstance(v, bool):
            return 0.0
        if isinstance(v, (int, float)):
            f = float(v)
            return 0.0 if math.isnan(f) else f
        s = str(v).strip()
        if not s or s.startswith("="):
            return 0.0
        s = s.replace(",", "").replace("₹", "").replace("%", "").strip()
        try:
            return float(s)
        except ValueError:
            return 0.0

    def _txt(row, key):
        idx = cm.get(key)
        if idx is None or idx >= len(row):
            return ""
        v = row[idx]
        if v is None or isinstance(v, bool):
            return ""
        s = str(v).strip()
        if s.endswith(".0") and s[:-2].isdigit():
            s = s[:-2]
        return "" if s.startswith("=") else s

    def _date(row, key):
        idx = cm.get(key)
        if idx is None or idx >= len(row):
            return datetime.now().date().isoformat()
        v = row[idx]
        if v is None:
            return datetime.now().date().isoformat()
        if hasattr(v, "strftime"):
            return v.strftime("%Y-%m-%d")
        s = str(v).strip()
        if not s or s.startswith("="):
            return datetime.now().date().isoformat()
        for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y", "%d.%m.%Y", "%d-%b-%Y"):
            try:
                from datetime import datetime as dt
                return dt.strptime(s[:10], fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        return s[:10]

    ERP_CATEGORY_MAP = {
        "erp_cogs":     "COGS",
        "erp_djcs":     "DJCS",
        "erp_equip":    "Equipment",
        "erp_manpower": "Manpower",
        "erp_vehicles": "Vehicles",
        "erp_others":   "Others",
    }

    updated = inserted = skipped = 0
    errors = []
    today = datetime.now().date().isoformat()
    now_iso = datetime.now().isoformat()

    # Pre-fetch existing projects from DB for fast in-memory matching
    with ENGINE.connect() as conn:
        all_db_projects = _rows_to_dicts(conn, "SELECT * FROM projects ORDER BY id ASC")

    proj_by_id = {}
    proj_by_co = {}
    proj_by_co_norm = {}
    proj_by_name = {}
    proj_by_name_norm = {}

    def _clean_key(val):
        return re.sub(r"[^a-z0-9]+", "", str(val or "").strip().lower())

    for p in all_db_projects:
        pid_val = p.get("id")
        if pid_val:
            proj_by_id[str(pid_val)] = p
        co_val = str(p.get("co_no") or "").strip()
        if co_val:
            proj_by_co[co_val.lower()] = p
            proj_by_co_norm[_clean_key(co_val)] = p
        name_val = str(p.get("project_name") or "").strip()
        if name_val:
            proj_by_name[name_val.lower()] = p
            proj_by_name_norm[_clean_key(name_val)] = p

    for row_num, row in enumerate(rows[hdr_idx + 1:], start=hdr_idx + 2):
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue

        raw_pid  = _txt(row, "project_id")
        raw_co   = _txt(row, "co_no")
        raw_name = _txt(row, "project_name")
        raw_pm   = _txt(row, "project_manager")
        raw_lob  = _txt(row, "lob")

        if not raw_pid and not raw_co and not raw_name:
            skipped += 1
            continue

        # Match project using multiple strategies
        proj = None
        if raw_co:
            proj = proj_by_co.get(raw_co.lower()) or proj_by_co_norm.get(_clean_key(raw_co))
        if not proj and raw_pid:
            proj = proj_by_id.get(raw_pid) or proj_by_co.get(raw_pid.lower()) or proj_by_co_norm.get(_clean_key(raw_pid))
        if not proj and raw_name:
            proj = proj_by_name.get(raw_name.lower()) or proj_by_name_norm.get(_clean_key(raw_name))

        # If project still not found, auto-create it so no project or manager expenditure is lost
        if not proj:
            try:
                co_to_use   = raw_co or raw_pid or (f"PRJ-{row_num}" if not raw_name else raw_name[:30])
                name_to_use = raw_name or f"Project {co_to_use}"
                lob_to_use  = raw_lob or ""
                pm_to_use   = raw_pm or ""

                with ENGINE.begin() as conn:
                    res = conn.execute(text("""
                        INSERT INTO projects (
                            co_no, project_name, hier_business_line, business_line,
                            project_manager, itd_cost, project_status, created_at
                        ) VALUES (
                            :co, :pname, :lob, :lob, :pm, 0, 'Active', :now
                        ) RETURNING id
                    """), {
                        "co": co_to_use, "pname": name_to_use, "lob": lob_to_use,
                        "pm": pm_to_use, "now": now_iso
                    })
                    new_pid = res.fetchone()[0]

                proj = {
                    "id": new_pid, "co_no": co_to_use, "project_name": name_to_use,
                    "hier_business_line": lob_to_use, "business_line": lob_to_use,
                    "project_manager": pm_to_use, "itd_cost": 0
                }
                # Register into fast lookup caches
                proj_by_id[str(new_pid)] = proj
                proj_by_co[co_to_use.lower()] = proj
                proj_by_co_norm[_clean_key(co_to_use)] = proj
                proj_by_name[name_to_use.lower()] = proj
                proj_by_name_norm[_clean_key(name_to_use)] = proj

                if pm_to_use:
                    try: add_ref_value("ref_project_managers", "name", pm_to_use)
                    except Exception: pass
                if lob_to_use:
                    try: add_ref_value("ref_lob", "name", lob_to_use)
                    except Exception: pass

            except Exception as e:
                errors.append(f"Row {row_num} (CO={raw_co or raw_pid!r}): Failed to auto-create project — {e}")
                skipped += 1
                continue

        pid = proj["id"]
        co_no_val = proj.get("co_no") or raw_co or raw_pid

        # Update PM / LOB if supplied and missing on project
        if raw_pm and not (proj.get("project_manager") or "").strip():
            try:
                with ENGINE.begin() as conn:
                    conn.execute(text("UPDATE projects SET project_manager = :pm WHERE id = :id"),
                                 {"pm": raw_pm, "id": pid})
                proj["project_manager"] = raw_pm
                add_ref_value("ref_project_managers", "name", raw_pm)
            except Exception: pass

        if raw_lob and not (proj.get("hier_business_line") or proj.get("business_line") or "").strip():
            try:
                with ENGINE.begin() as conn:
                    conn.execute(text("UPDATE projects SET hier_business_line = :l, business_line = :l WHERE id = :id"),
                                 {"l": raw_lob, "id": pid})
                proj["hier_business_line"] = raw_lob
                add_ref_value("ref_lob", "name", raw_lob)
            except Exception: pass

        lob_val         = raw_lob or proj.get("hier_business_line") or proj.get("business_line") or ""
        data_as_of      = _date(row, "data_as_of")      or today
        erp_update_date = _date(row, "erp_update_date") or today

        itd_cogs     = _num(row, "itd_cogs")
        itd_djcs     = _num(row, "itd_djcs")
        itd_equip    = _num(row, "itd_equip")
        itd_manpower = _num(row, "itd_manpower")
        itd_vehicles = _num(row, "itd_vehicles")
        itd_others   = _num(row, "itd_others")
        total_itd_col= _num(row, "total_itd")

        itd_sum = itd_cogs + itd_djcs + itd_equip + itd_manpower + itd_vehicles + itd_others
        if itd_sum == 0 and total_itd_col > 0:
            itd_others = total_itd_col
            itd_sum    = total_itd_col

        erp_cogs     = _num(row, "erp_cogs")
        erp_djcs     = _num(row, "erp_djcs")
        erp_equip    = _num(row, "erp_equip")
        erp_manpower = _num(row, "erp_manpower")
        erp_vehicles = _num(row, "erp_vehicles")
        erp_others   = _num(row, "erp_others")
        total_erp_col= _num(row, "total_erp")

        erp_sum = erp_cogs + erp_djcs + erp_equip + erp_manpower + erp_vehicles + erp_others
        if erp_sum == 0 and total_erp_col > 0:
            erp_others = total_erp_col
            erp_sum    = total_erp_col
        if erp_sum == 0 and itd_sum > 0:
            erp_cogs     = itd_cogs
            erp_djcs     = itd_djcs
            erp_equip    = itd_equip
            erp_manpower = itd_manpower
            erp_vehicles = itd_vehicles
            erp_others   = itd_others
            erp_sum      = itd_sum

        erp_vals = {
            "erp_cogs":     erp_cogs,
            "erp_djcs":     erp_djcs,
            "erp_equip":    erp_equip,
            "erp_manpower": erp_manpower,
            "erp_vehicles": erp_vehicles,
            "erp_others":   erp_others,
        }

        exp_fields = {
            "co_no":           co_no_val,
            "lob":             lob_val,
            "data_as_of":      data_as_of,
            "erp_update_date": erp_update_date,
            "itd_cogs":        itd_cogs,
            "itd_djcs":        itd_djcs,
            "itd_equip":       itd_equip,
            "itd_manpower":    itd_manpower,
            "itd_vehicles":    itd_vehicles,
            "itd_others":      itd_others,
            "erp_cogs":        erp_cogs,
            "erp_djcs":        erp_djcs,
            "erp_equip":       erp_equip,
            "erp_manpower":    erp_manpower,
            "erp_vehicles":    erp_vehicles,
            "erp_others":      erp_others,
        }

        try:
            with ENGINE.begin() as conn:
                # ── STEP 1: Replace erp_updates rows for this project ──
                conn.execute(text(
                    "DELETE FROM erp_updates WHERE co_no = :co OR project_id = :pid"
                ), {"co": co_no_val, "pid": pid})

                # ── STEP 2: Insert one erp_updates row per category ──
                for erp_key, category in ERP_CATEGORY_MAP.items():
                    amt = erp_vals[erp_key]
                    if amt:
                        conn.execute(text("""
                            INSERT INTO erp_updates
                                (project_id, co_no, category, amount, created_by, created_at)
                            VALUES
                                (:pid, :co, :cat, :amt, :by, :ts)
                        """), {
                            "pid": pid,
                            "co":  co_no_val,
                            "cat": category,
                            "amt": amt,
                            "by":  "excel_upload",
                            "ts":  now_iso,
                        })

                # ── STEP 3: Upsert expenditures row ──
                existing = conn.execute(text(
                    "SELECT id FROM expenditures WHERE co_no=:c OR project_id=:pid ORDER BY id DESC LIMIT 1"
                ), {"c": co_no_val, "pid": pid}).fetchone()

                exp_fields["project_id"] = pid

                if existing:
                    conn.execute(text("""
                        UPDATE expenditures SET
                            project_id=:project_id,
                            lob=:lob,
                            data_as_of=:data_as_of,
                            erp_update_date=:erp_update_date,
                            itd_cogs=:itd_cogs,
                            itd_djcs=:itd_djcs,
                            itd_equip=:itd_equip,
                            itd_manpower=:itd_manpower,
                            itd_vehicles=:itd_vehicles,
                            itd_others=:itd_others,
                            erp_cogs=:erp_cogs,
                            erp_djcs=:erp_djcs,
                            erp_equip=:erp_equip,
                            erp_manpower=:erp_manpower,
                            erp_vehicles=:erp_vehicles,
                            erp_others=:erp_others
                        WHERE id=:eid
                    """), {**exp_fields, "eid": existing[0]})
                    updated += 1
                else:
                    exp_fields["created_at"] = now_iso
                    conn.execute(text("""
                        INSERT INTO expenditures (
                            project_id, co_no, lob, data_as_of, erp_update_date,
                            itd_cogs, itd_djcs, itd_equip, itd_manpower,
                            itd_vehicles, itd_others,
                            erp_cogs, erp_djcs, erp_equip, erp_manpower,
                            erp_vehicles, erp_others, created_at
                        ) VALUES (
                            :project_id, :co_no, :lob, :data_as_of, :erp_update_date,
                            :itd_cogs, :itd_djcs, :itd_equip, :itd_manpower,
                            :itd_vehicles, :itd_others,
                            :erp_cogs, :erp_djcs, :erp_equip, :erp_manpower,
                            :erp_vehicles, :erp_others, :created_at
                        )
                    """), exp_fields)
                    inserted += 1

                # ── STEP 4: Sync itd_cost on the project row ──
                live_total_res = conn.execute(text(
                    "SELECT COALESCE(SUM(amount),0) FROM erp_updates WHERE project_id=:pid OR co_no=:co"
                ), {"pid": pid, "co": co_no_val}).fetchone()
                live_total = float(live_total_res[0]) if live_total_res else 0.0

                proj_sheet_itd = projects_sheet_itd.get(co_no_val.lower(), 0.0)
                cur_itd_res = conn.execute(text("SELECT itd_cost FROM projects WHERE id=:id"), {"id": pid}).fetchone()
                existing_itd = float(cur_itd_res[0] or 0.0) if cur_itd_res else 0.0

                final_itd = max(total_itd_col, proj_sheet_itd, itd_sum, existing_itd)
                if final_itd == 0:
                    final_itd = live_total

                diff = final_itd - live_total
                if diff > 0.001:
                    conn.execute(text("""
                        INSERT INTO erp_updates
                            (project_id, co_no, category, amount, created_by, created_at)
                        VALUES (:pid, :co, 'Others', :amt, 'excel_upload', :ts)
                    """), {
                        "pid": pid, "co": co_no_val, "amt": diff, "ts": now_iso
                    })

                conn.execute(text(
                    "UPDATE projects SET itd_cost=:v WHERE id=:id"
                ), {"v": final_itd, "id": pid})

        except Exception as e:
            errors.append(f"Row {row_num} (CO={co_no_val!r}): {e}")

    # Ensure any projects present in Projects sheet but not in the Expenditures rows are updated
    if projects_sheet_itd:
        with ENGINE.begin() as conn:
            for p_co, p_itd in projects_sheet_itd.items():
                p_obj = proj_by_co.get(p_co) or proj_by_co_norm.get(_clean_key(p_co))
                if p_obj:
                    p_id = p_obj["id"]
                    cur = conn.execute(text("SELECT itd_cost FROM projects WHERE id=:id"), {"id": p_id}).fetchone()
                    cur_v = float(cur[0] or 0.0) if cur else 0.0
                    if p_itd > cur_v:
                        conn.execute(text("UPDATE projects SET itd_cost=:v WHERE id=:id"), {"v": p_itd, "id": p_id})
                        conn.execute(text("""
                            INSERT INTO erp_updates (project_id, co_no, category, amount, created_by, created_at)
                            VALUES (:pid, :co, 'Others', :amt, 'excel_upload', :ts)
                        """), {"pid": p_id, "co": p_obj.get("co_no") or p_co, "amt": p_itd, "ts": now_iso})

    msg = (f"✅ Expenditure upload complete — "
           f"{inserted} new records inserted, {updated} existing records updated"
           + (f", {skipped} rows skipped" if skipped else "") + ".")

    if errors:
        err_detail = "<br>".join(errors[:10])
        return (f'<div class="alert ok">{msg}</div>'
                f'<div class="alert err">⚠ Some rows had issues:<br>{err_detail}</div>')
    return f'<div class="alert ok">{msg}</div>'
def expenditures(alert_html="", prefill=None, username="Admin") -> str:
    p     = prefill or {}
    today = date.today().isoformat()

    lob_list   = get_visible_lob_list()
    pm_list    = get_project_managers()
    exp_cats   = get_expense_categories()
    cat_colors = get_expense_category_colors()
    proj_rows  = fetch_projects_for_select()

    # co_no (e.g. "620_7") is the linking key for Expenditures now, not the internal id.
    sel_co   = p.get("co_no","") or p.get("project_id","")   # project_id kept as fallback for old links
    sel_lob  = p.get("lob","")
    sel_pm   = p.get("pm","")

    # ── Filter projects by LOB / PM if selected ──────────────────────────────
    filtered_rows = proj_rows
    if sel_lob:
        filtered_rows = [r for r in filtered_rows
                         if str(r[6] or "").strip().lower() == sel_lob.strip().lower()
                         or str(r[9] or "").strip().lower() == sel_lob.strip().lower()]
    if sel_pm:
        filtered_rows = [r for r in filtered_rows
                         if str(r[10] or "").strip().lower() == sel_pm.strip().lower()]

    proj_opts = "".join(
        f'<option value="{r[1] or ""}" {"selected" if str(sel_co)==str(r[1] or "") else ""}>'
        f'{r[1] or "(no CO.No)"} | {r[2] or ""}</option>'
        for r in filtered_rows if r[1]
    ) or '<option value="">— No projects for this filter —</option>'

    sel_proj = next((r for r in proj_rows if str(r[1] or "")==str(sel_co)), None)

    meta_client = sel_proj[3] if sel_proj else "—"
    meta_ordno  = sel_proj[4] if sel_proj else "—"
    sel_rev_ov  = ((_n(sel_proj[8]) or (_n(sel_proj[5]) + _n(sel_proj[7]))) if sel_proj else 0)
    meta_ordval = f"₹ {sel_rev_ov:,.0f}" if sel_rev_ov else "—"
    meta_lob    = ((sel_proj[9] or sel_proj[6]) if sel_proj else "—")

    lob_opts = "".join(
        f'<option value="{lob}" {"selected" if sel_lob == lob else ""}>{lob}</option>'
        for lob in lob_list
    )
    pm_opts = "".join(
        f'<option value="{pmn}" {"selected" if sel_pm == pmn else ""}>{pmn}</option>'
        for pmn in pm_list
    )

    existing_exp = {}
    if sel_co:
        existing_exp = fetch_expenditure_for_co_no(sel_co)

    erp_totals = fetch_erp_totals(sel_co) if sel_co else {}

    itd_defaults = {
        "COGS":      "erp_cogs",
        "DJCS":      "erp_djcs",
        "Equipment": "erp_equip",
        "Manpower":  "erp_manpower",
        "Vehicles":  "erp_vehicles",
        "Others":    "erp_others",
    }

    # ── ERP Reconciliation columns ───────────────────────────────────────────
    # cumulative_val  = total from erp_updates  (shown in the blue card)
    # stored_erp_val  = snapshot in expenditures (shown in input as readonly)
    erp_cols_html = ""
    for cat in exp_cats:
        col_key           = f"erp_{cat.lower().replace(' ','_').replace('-','_')}"
        legacy            = itd_defaults.get(cat, col_key)
        stored_erp_key    = legacy                          # e.g. "erp_cogs"
        cumulative_val    = erp_totals.get(cat, 0)         # from erp_updates SUM
        stored_erp_val    = _n(existing_exp.get(stored_erp_key, 0))  # from expenditures

        # The display card shows the cumulative total from erp_updates
        # The input shows the stored snapshot — pre-filled and readonly after upload
        erp_cols_html += f"""
          <div style="min-width:0">
            <!-- Cumulative total card -->
            <div style="background:{C['surface_low']};border:1px solid {C['outline']};border-radius:8px;
                        padding:12px 14px;margin-bottom:10px">
              <div style="font-size:9px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;
                          color:{C['on_var']};margin-bottom:6px">{cat.upper()}</div>
              <div style="font-size:16px;font-weight:800;color:{C['dark']};letter-spacing:-.02em">
                ₹ {_n(cumulative_val):,.0f}
              </div>
               <div style="font-size:10px;color:{C['on_var']};margin-top:3px;opacity:.7">Cumulative Total Expenditure</div>
            </div>
            <!-- ERP input — always starts at 0 for new delta entry -->
            <div style="font-size:9px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;
                        color:{C['on_var']};margin-bottom:6px">ERP {cat.upper()}</div>
            <div style="position:relative">
              <span style="position:absolute;left:10px;top:50%;transform:translateY(-50%);
                          font-size:12px;color:{C['on_var']};pointer-events:none">₹</span>
              <input type="number" id="{legacy}" name="{legacy}"
                    value="0"
                    step="0.01"
                    data-existing="{_n(cumulative_val)}"
                    oninput="calcErpTotal()"
                    style="width:100%;padding:9px 9px 9px 24px;border:1px solid {C['outline']};
                            border-radius:6px;font-size:13px;font-weight:600;background:#fff;
                            outline:none;box-sizing:border-box">
            </div>
          </div>"""

    # ── Donut chart — use erp_totals (cumulative) not existing_exp snapshot ──
    donut_segs = []
    palette    = ["#003f88","#d00000","#ffba08","#3a86ff","#06a77d","#7209b7","#fb5607","#1a1a2e"]

    erp_chart_vals = {cat: _n(erp_totals.get(cat, 0)) for cat in exp_cats}
    total_itd_chart = sum(erp_chart_vals.values())

    # Fallback to expenditures snapshot if erp_updates is empty
    if total_itd_chart == 0:
        legacy_map_donut = {
            "COGS":      "erp_cogs",
            "DJCS":      "erp_djcs",
            "Equipment": "erp_equip",
            "Manpower":  "erp_manpower",
            "Vehicles":  "erp_vehicles",
            "Others":    "erp_others",
        }
        erp_chart_vals = {
            cat: _n(existing_exp.get(
                legacy_map_donut.get(cat, f"erp_{cat.lower().replace(' ','_')}"), 0
            ))
            for cat in exp_cats
        }
        total_itd_chart = sum(erp_chart_vals.values())

    total_itd = total_itd_chart or 1
    for i, cat in enumerate(exp_cats):
        bg        = palette[i % len(palette)]
        val       = erp_chart_vals[cat]
        pct_exact = (val / total_itd * 100) if total_itd else (100 / max(len(exp_cats), 1))
        pct_disp  = round(pct_exact)
        donut_segs.append((cat, bg, pct_exact, pct_disp, val))

    if total_itd_chart >= 10_00_000:
        total_itd_label = f"₹{total_itd_chart/10_00_000:.1f}M"
    elif total_itd_chart >= 1_00_000:
        total_itd_label = f"₹{total_itd_chart/1_00_000:.1f}L"
    else:
        total_itd_label = f"₹{total_itd_chart:,.0f}"

    # ── Breakup of expenditure by head — line-item entries logged in expense_entries,
    #    PLUS the ERP Cost Reconciliation deltas logged in erp_updates (so an ERP
    #    update shows up here too, with its amount and date, not just as a silent
    #    change to the cumulative total). ──
    entries_by_cat: dict = {}
    if sel_co:
        for e in fetch_expense_entries(sel_co):
            entries_by_cat.setdefault(e.get("category") or "Uncategorized", []).append(e)
        for u in fetch_erp_update_entries(sel_co):
            entries_by_cat.setdefault(u.get("category") or "Uncategorized", []).append({
                "id":          u.get("id"),
                "description": "ERP Reconciliation Update",
                "quantity":    None,
                "unit_cost":   None,
                "line_total":  u.get("amount"),
                "created_at":  u.get("created_at"),
                "source":      "erp",
            })

    import json as _json_exp
    import html as _html_esc   # local alias — `html` is reused below as the page-string var

    def _iso(v):
        try:
            return str(v)[:10] if v else ""
        except Exception:
            return ""

    def _entry_sort_key(e):
        return _iso(e.get("created_at")) or ""

    head_entries_payload = {
        cat: [
            {
                "id":          e.get("id"),
                "description": e.get("description") or "",
                "quantity":    (None if e.get("source") == "erp" else _n(e.get("quantity"))),
                "unit_cost":   (None if e.get("source") == "erp" else _n(e.get("unit_cost"))),
                "line_total":  _n(e.get("line_total")),
                "created_at":  _iso(e.get("created_at")),
                "source":      e.get("source", "manual"),
            }
            for e in sorted(entries_by_cat.get(cat, []), key=_entry_sort_key, reverse=True)
        ]
        for cat in exp_cats
    }
    head_entries_json = _json_exp.dumps(head_entries_payload).replace("</", "<\\/")

    breakup_cards_html = ""
    for i, cat in enumerate(exp_cats) if sel_co else []:
        bg      = palette[i % len(palette)]
        amt     = _n(erp_totals.get(cat, 0))
        entries = entries_by_cat.get(cat, [])
        cnt     = len(entries)
        cat_js  = _html_esc.escape(cat).replace("'", "\\'")
        breakup_cards_html += f"""
          <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;
                      padding:18px 20px;display:flex;flex-direction:column;gap:10px">
            <div style="display:flex;align-items:center;gap:8px">
              <span style="width:10px;height:10px;border-radius:50%;background:{bg};flex-shrink:0;
                          display:inline-block"></span>
              <span style="font-size:11px;font-weight:700;text-transform:uppercase;
                          letter-spacing:.06em;color:{C['on_var']}">{_html_esc.escape(cat)}</span>
            </div>
            <div style="font-size:19px;font-weight:800;color:{C['dark']};letter-spacing:-.02em">
              ₹ {amt:,.0f}
            </div>
            <div style="font-size:11px;color:{C['on_var']};opacity:.75">
              {cnt} logged {"entry" if cnt == 1 else "entries"}
            </div>
            <button type="button" onclick="openHeadDetail('{cat_js}')"
                    style="margin-top:6px;padding:8px 12px;border:1px solid {C['outline']};
                           border-radius:6px;background:{C['surface_low']};color:{C['dark']};
                           font-size:10px;font-weight:700;text-transform:uppercase;
                           letter-spacing:.06em;cursor:pointer;text-align:center;
                           transition:background .15s"
                    onmouseover="this.style.background='{C['surface_high']}'"
                    onmouseout="this.style.background='{C['surface_low']}'">
              View Details
            </button>
          </div>"""

    r_val = 15.9155
    sw    = 14/140*36
    cx, cy = 18, 18
    import math

    def _arc_point(pct):
        ang = math.radians(-90 + (pct / 100.0) * 360)
        return cx + r_val * math.cos(ang), cy + r_val * math.sin(ang)

    arc_parts = [f'<circle cx="{cx}" cy="{cy}" r="{r_val}" fill="none" stroke="#f2f4f6" stroke-width="{sw:.2f}"/>']
    cur = 0.0
    for label, colour, pct_exact, pct_disp, valamt in donut_segs:
        if pct_exact <= 0.0001:
            continue
        start_pct, end_pct = cur, cur + pct_exact
        x1, y1 = _arc_point(start_pct)
        x2, y2 = _arc_point(min(end_pct, 99.999))
        large_arc = 1 if pct_exact > 50 else 0
        arc_parts.append(
            f'<path d="M {x1:.4f} {y1:.4f} A {r_val} {r_val} 0 {large_arc} 1 {x2:.4f} {y2:.4f}" '
            f'fill="none" stroke="{colour}" stroke-width="{sw:.2f}" stroke-linecap="butt" '
            f'style="cursor:pointer" '
            f'onmousemove="showDonutTip(event,\'{label}\',\'{_n(valamt):,.0f}\',\'{pct_disp}\')" '
            f'onmouseleave="hideDonutTip()"/>'
        )
        cur = end_pct
    donut_svg = (
        f'<svg width="160" height="160" viewBox="0 0 36 36">'
        + "".join(arc_parts)
        + '</svg>'
    )

    legend_items = "".join(
        f'<div style="display:flex;align-items:center;justify-content:space-between;'
        f'padding:7px 0;border-bottom:1px solid {C["outline"]}">'
        f'<div style="display:flex;align-items:center;gap:8px">'
        f'<span style="width:10px;height:10px;border-radius:50%;background:{col};'
        f'flex-shrink:0;display:inline-block"></span>'
        f'<span style="font-size:13px;color:{C["on_var"]}">{lbl}</span>'
        f'</div>'
        f'<span style="font-size:13px;font-weight:700;color:{C["dark"]}">{pct_disp}%</span>'
        f'</div>'
        for lbl, col, pct_exact, pct_disp, valamt in donut_segs
    )

    data_as_of = existing_exp.get("data_as_of") or p.get("data_as_of", today)
    erp_update = existing_exp.get("erp_update_date") or p.get("erp_update_date", today)
    cat_opts   = "".join(f'<option value="{c}">{c}</option>' for c in exp_cats)
    legacy_map = {
        "COGS":      "erp_cogs",
        "DJCS":      "erp_djcs",
        "Equipment": "erp_equip",
        "Manpower":  "erp_manpower",
        "Vehicles":  "erp_vehicles",
        "Others":    "erp_others",
    }
    erp_field_ids = [legacy_map.get(c, f"erp_{c.lower().replace(' ','_')}") for c in exp_cats]
    erp_ids_js    = "['" + "','".join(erp_field_ids) + "']"

    css = f"""<style>
*{{box-sizing:border-box}}
.exp-title{{font-size:24px;font-weight:800;color:{C['dark']};letter-spacing:-.02em}}

/* ── Filter bar — no button, dropdowns side by side ── */
.sel-bar{{
  background:#fff;border:1px solid {C['outline']};border-radius:10px;
  padding:16px 20px;display:flex;align-items:flex-end;gap:16px;
  margin-bottom:20px;
}}
.sel-bar .field{{flex:1;min-width:0}}
.sel-bar .field label{{
  font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
  color:{C['on_var']};display:block;margin-bottom:6px
}}
.sel-bar .field select{{
  width:100%;padding:9px 12px;border:1px solid {C['outline']};border-radius:6px;
  font-size:13px;background:#fff;outline:none;color:{C['dark']};cursor:pointer;
  transition:border-color .15s;
}}
.sel-bar .field select:focus{{border-color:{C['dark']}}}

.meta-strip{{display:grid;grid-template-columns:repeat(4,1fr);background:#fff;
             border:1px solid {C['outline']};border-radius:10px;overflow:hidden}}
.meta-cell{{padding:16px 20px;border-right:1px solid {C['outline']}}}
.meta-cell:last-child{{border-right:none}}
.mlbl{{font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
       color:{C['on_var']};margin-bottom:5px}}
.mval{{font-size:15px;font-weight:700;color:{C['dark']}}}
.side-panel{{position:fixed;inset-y:0;right:0;width:400px;background:#fff;
             border-left:1px solid {C['outline']};box-shadow:-4px 0 32px rgba(0,0,0,.10);
             z-index:300;transform:translateX(100%);transition:transform .3s ease;
             display:flex;flex-direction:column}}
.side-panel.open{{transform:translateX(0)}}
.sp-hdr{{padding:20px 24px;border-bottom:1px solid {C['outline']};
         background:{C['surface_low']};display:flex;align-items:center;justify-content:space-between}}
.sp-body{{flex:1;padding:24px;overflow-y:auto}}
.sp-fld{{margin-bottom:16px}}
.sp-fld label{{font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
               color:{C['on_var']};display:block;margin-bottom:6px}}
.sp-fld input,.sp-fld select{{width:100%;padding:9px 12px;border:1px solid {C['outline']};
                               border-radius:6px;font-size:13px;background:#fff;
                               outline:none;color:{C['dark']}}}
.sp-g2{{display:grid;grid-template-columns:1fr 1fr;gap:12px}}
.sp-foot{{padding:16px 24px;border-top:1px solid {C['outline']};
          background:{C['surface_low']};display:flex;gap:10px}}
.overlay{{position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:299;display:none}}
.overlay.show{{display:block}}
.erp-card{{background:#fff;border:1px solid {C['outline']};border-radius:10px;overflow:hidden}}
.erp-hdr{{background:{C['surface_low']};padding:18px 28px;border-bottom:1px solid {C['outline']};
          display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:16px}}
.data-tbl{{width:100%;border-collapse:collapse;font-size:13px}}
.data-tbl thead tr{{background:{C['surface_low']};border-bottom:1px solid {C['outline']}}}
.data-tbl th{{padding:10px 16px;font-size:9px;font-weight:700;text-transform:uppercase;
              letter-spacing:.07em;color:{C['on_var']};text-align:left}}
.data-tbl tbody tr{{border-bottom:1px solid {C['outline']}}}
.data-tbl tbody tr:last-child{{border-bottom:none}}
.data-tbl tbody tr:hover{{background:{C['surface_low']}}}
.exp-upload-zone{{
  border:2px dashed {C['outline']};border-radius:12px;padding:22px 28px;
  background:#fff;text-align:center;cursor:pointer;margin-bottom:20px;
  transition:border-color .2s,background .2s;display:flex;align-items:center;
  gap:20px;flex-wrap:wrap;justify-content:center;
}}
.exp-upload-zone:hover{{border-color:{C['dark']};background:#f0f4ff}}
.exp-upload-zone .uz-icon{{
  width:48px;height:48px;background:{C['dark']};border-radius:50%;
  display:flex;align-items:center;justify-content:center;
  font-size:22px;color:#fff;flex-shrink:0;
}}
.exp-upload-zone .uz-text{{text-align:left}}
.exp-upload-zone .uz-text h4{{font-size:15px;font-weight:700;color:{C['dark']};margin-bottom:4px}}
.exp-upload-zone .uz-text p{{font-size:12px;color:#6b7280}}
.exp-upload-zone .uz-badge{{
  display:inline-flex;gap:10px;background:{C['surface_high']};
  padding:5px 14px;border-radius:20px;font-size:11px;font-weight:700;
  letter-spacing:.07em;color:{C['on_var']};white-space:nowrap;
}}
</style>"""

    js = f"""<script>
/* ── Auto-submit on LOB change: repopulate project list ── */
function onLobChange(sel) {{
  document.getElementById('exp-filter-form').submit();
}}

/* ── Auto-submit on project change: load that project's data ── */
function onProjectChange(sel) {{
  if (sel.value) {{
    document.getElementById('exp-filter-form').submit();
  }}
}}

function openPanel(){{
  document.getElementById('sp-cat').value='{exp_cats[0] if exp_cats else ""}';
  document.getElementById('sp-desc').value='';
  document.getElementById('sp-qty').value='';
  document.getElementById('sp-ucost').value='';
  document.getElementById('sp-total').textContent='₹ 0.00';
  document.getElementById('side-panel').classList.add('open');
  document.getElementById('overlay').classList.add('show');
}}
function closePanel(){{
  document.getElementById('side-panel').classList.remove('open');
  document.getElementById('overlay').classList.remove('show');
}}
document.addEventListener('keydown',function(e){{if(e.key==='Escape')closePanel();}});
function calcTotal(){{
  var q=parseFloat(document.getElementById('sp-qty').value)||0;
  var u=parseFloat(document.getElementById('sp-ucost').value)||0;
  document.getElementById('sp-total').textContent='₹ '+(q*u).toLocaleString('en-IN',{{minimumFractionDigits:2}});
}}
function calcErpTotal(){{
  var ids={erp_ids_js};
  var total=0;
  ids.forEach(function(id){{
    var el=document.getElementById(id);
    if(el) total += (parseFloat(el.dataset.existing)||0) + (parseFloat(el.value)||0);
  }});
  document.getElementById('erp-total').value=total.toLocaleString('en-IN',{{minimumFractionDigits:2}});
}}
/* ── Expenditure breakup by head — detail modal ── */
var HEAD_ENTRIES = {head_entries_json};
function fmtMoney(n){{
  return '₹ ' + (Number(n)||0).toLocaleString('en-IN',{{minimumFractionDigits:2}});
}}
function openHeadDetail(cat){{
  document.getElementById('head-modal-title').textContent = cat;
  var rows = HEAD_ENTRIES[cat] || [];
  var body = document.getElementById('head-modal-body');
  if(!rows.length){{
    body.innerHTML = '<div style="text-align:center;padding:34px 0;color:#9ca3af;font-size:13px">'
                    + 'No line-item entries logged for this head yet.</div>';
  }} else {{
    var out = '<table class="data-tbl"><thead><tr>'
            + '<th>Description</th><th>Qty</th><th>Unit Cost</th><th>Line Total</th><th>Date</th><th></th>'
            + '</tr></thead><tbody>';
    var total = 0;
    rows.forEach(function(r){{
      total += (Number(r.line_total) || 0);
      var isErp = r.source === 'erp';
      var descHtml = (r.description || '—') + (isErp
        ? ' <span style="display:inline-block;margin-left:6px;padding:1px 7px;border-radius:10px;'
          + 'background:#e0f2fe;color:#0369a1;font-size:9px;font-weight:700;letter-spacing:.04em;'
          + 'text-transform:uppercase;vertical-align:middle">ERP</span>'
        : '');
      var qtyHtml  = (r.quantity  === null || r.quantity  === undefined) ? '—' : r.quantity;
      var costHtml = (r.unit_cost === null || r.unit_cost === undefined) ? '—' : fmtMoney(r.unit_cost);
      var actionHtml = isErp
        ? '<span style="color:#9ca3af;font-size:11px">—</span>'
        : '<button type="button" onclick="deleteHeadEntry(' + r.id + ')" '
          + 'style="background:none;border:none;color:{C["error"]};font-size:11px;'
          + 'font-weight:700;cursor:pointer;text-decoration:underline">Delete</button>';
      out += '<tr>'
           + '<td style="padding:10px 14px">' + descHtml + '</td>'
           + '<td style="padding:10px 14px">' + qtyHtml + '</td>'
           + '<td style="padding:10px 14px">' + costHtml + '</td>'
           + '<td style="padding:10px 14px;font-weight:700">' + fmtMoney(r.line_total) + '</td>'
           + '<td style="padding:10px 14px;color:#6b7280">' + (r.created_at || '—') + '</td>'
           + '<td style="padding:10px 14px;text-align:right">' + actionHtml + '</td>'
           + '</tr>';
    }});
    out += '</tbody></table>'
         + '<div style="margin-top:16px;padding-top:14px;border-top:1px solid {C["outline"]};'
         + 'display:flex;justify-content:space-between;align-items:center">'
         + '<span style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;'
         + 'color:{C["on_var"]}">Head Total</span>'
         + '<span style="font-size:16px;font-weight:800;color:{C["dark"]}">' + fmtMoney(total) + '</span>'
         + '</div>';
    body.innerHTML = out;
  }}
  document.getElementById('head-overlay').classList.add('show');
  var modal = document.getElementById('head-detail-modal');
  modal.style.display = 'flex';
  requestAnimationFrame(function(){{
    modal.style.opacity = '1';
    modal.style.transform = 'translate(-50%,-50%) scale(1)';
  }});
}}
function closeHeadDetail(){{
  document.getElementById('head-overlay').classList.remove('show');
  var modal = document.getElementById('head-detail-modal');
  modal.style.opacity = '0';
  modal.style.transform = 'translate(-50%,-50%) scale(.96)';
  setTimeout(function(){{ modal.style.display = 'none'; }}, 150);
}}
function deleteHeadEntry(id){{
  if(!confirm('Delete this expense entry? This cannot be undone.')) return;
  document.getElementById('del-entry-id').value = id;
  document.getElementById('del-entry-form').submit();
}}
document.addEventListener('keydown', function(e){{ if(e.key === 'Escape') closeHeadDetail(); }});

function showDonutTip(e, label, value, pct){{
  var tip  = document.getElementById('donut-tooltip');
  var wrap = document.getElementById('donut-wrap');
  var rect = wrap.getBoundingClientRect();
  tip.innerHTML = '<div style="font-weight:700;margin-bottom:2px">'+label+'</div>'
                + '<div>₹ '+value+' ('+pct+'%)</div>';
  tip.style.left = (e.clientX - rect.left) + 'px';
  tip.style.top  = (e.clientY - rect.top - 8) + 'px';
  tip.style.opacity = '1';
}}
function hideDonutTip(){{
  document.getElementById('donut-tooltip').style.opacity = '0';
}}
window.addEventListener('load', calcErpTotal);
</script>"""

    html  = _html_head("Expenditure Updates") + css + "<body>"
    html += _sidebar("expenditures") + _topbar("EXPENDITURE & TOTAL EXPENDITURE BREAKDOWN", username)
    html += f"""
<div class="main"><div class="canvas" style="max-width:1200px">

  <!-- Page header -->
  <div style="border-bottom:1px solid {C['outline']};padding-bottom:20px;margin-bottom:24px;
              display:flex;justify-content:space-between;align-items:flex-end">
    <div>
      <h2 class="exp-title">Expenditure Updates</h2>
      <p style="font-size:13px;color:#6b7280;margin-top:4px">
        Comprehensive analysis and ERP reconciliation for current projects.
      </p>
    </div>
    <a href="/export_excel/expenditures"
       style="display:inline-flex;align-items:center;gap:6px;padding:8px 16px;
              background:{C['dark']};color:#fff;border-radius:7px;font-size:12px;
              font-weight:700;text-decoration:none;letter-spacing:.04em"
    >⬇ Export Excel</a>
  </div>

  {alert_html}

<!-- ── EXCEL UPLOAD ZONE ── -->
  <form method="post" action="/upload_expenditure" enctype="multipart/form-data">
    <div class="exp-upload-zone" onclick="document.getElementById('exp-fi').click()">
      <div class="uz-icon">📤</div>
      <div class="uz-text">
        <h4>Bulk Upload via Excel</h4>
        <p>Upload your expenditure template — columns will be matched by project_id / CO.No
           and saved directly to the database.</p>
      </div>
      <div>
        <div class="uz-badge">XLSX &nbsp;/&nbsp; MAX 25MB</div>
        <p style="font-size:11px;color:#9ca3af;margin-top:6px;text-align:center">
          Click or drag to upload
        </p>
      </div>
      <input id="exp-fi" type="file" name="expenditure_file" accept=".xlsx"
             style="display:none" onchange="this.form.submit()">
    </div>
  </form>

  <!-- ── FILTER BAR — no button, auto-submits on change ── -->
  <form method="post" action="/expenditures_retrieve" id="exp-filter-form">
    <div class="sel-bar">
      <div class="field">
        <label>Line of Business (LOB)</label>
        <select name="lob" onchange="onLobChange(this)">
          <option value="">All LOB</option>
          {lob_opts}
        </select>
      </div>
      <div class="field">
        <label>Project Manager (PM)</label>
        <select name="pm" onchange="onLobChange(this)">
          <option value="">All PM</option>
          {pm_opts}
        </select>
      </div>
      <div class="field">
        <label>Active Project</label>
        <select name="co_no" onchange="onProjectChange(this)">
          <option value="">— Select a project —</option>
          {proj_opts}
        </select>
      </div>
    </div>
  </form>

  <!-- Project meta strip — only shown when a project is selected -->
  {'<div class="meta-strip" style="margin-bottom:28px">' +
    f'<div class="meta-cell"><div class="mlbl">Client Name</div><div class="mval">{meta_client}</div></div>' +
    f'<div class="meta-cell"><div class="mlbl">Order Number</div><div class="mval">{meta_ordno}</div></div>' +
    f'<div class="meta-cell"><div class="mlbl">Order Value</div><div class="mval">{meta_ordval}</div></div>' +
    f'<div class="meta-cell"><div class="mlbl">LOB</div><div class="mval">{meta_lob}</div></div>' +
   '</div>'
   if sel_co else
   f'<div style="background:{C["surface_low"]};border:1px solid {C["outline"]};border-radius:10px;'
   f'padding:24px;text-align:center;color:{C["on_var"]};font-size:13px;margin-bottom:28px">'
   f'← Select a project above to view its expenditure data</div>'
  }

  <!-- Donut chart + legend -->
  <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;
              padding:28px 32px;display:flex;align-items:center;gap:40px;margin-bottom:28px">
    <div style="position:relative;width:160px;height:160px;flex-shrink:0" id="donut-wrap">
      {donut_svg}
      <div style="position:absolute;inset:0;display:flex;flex-direction:column;
                  align-items:center;justify-content:center;text-align:center;
                  gap:2px;pointer-events:none">
        <span style="font-size:18px;font-weight:800;color:{C['dark']};
                     letter-spacing:-.03em">{total_itd_label}</span>
        <span style="font-size:9px;font-weight:700;text-transform:uppercase;
                     color:{C['on_var']};letter-spacing:.06em">Total Expenditure</span>
      </div>
      <div id="donut-tooltip"
           style="position:absolute;pointer-events:none;opacity:0;transition:opacity .12s;
                  background:{C['dark']};color:#fff;font-size:11px;padding:6px 10px;
                  border-radius:6px;white-space:nowrap;z-index:50;transform:translateY(-100%)">
      </div>
    </div>
    <div style="flex:1;display:grid;grid-template-columns:1fr 1fr;gap:0 40px">
      {legend_items}
    </div>
  </div>

  <!-- Breakup of Expenditure by Head -->
  <div style="margin-bottom:28px">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px">
      <h3 style="font-size:15px;font-weight:800;color:{C['dark']}">Expenditure Breakup by Head</h3>
      <span style="font-size:11px;color:{C['on_var']}">Click a head to view its logged line items</span>
    </div>
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:16px">
      {breakup_cards_html or f'<div style="grid-column:1/-1;background:{C["surface_low"]};border:1px solid {C["outline"]};border-radius:10px;padding:20px;text-align:center;color:{C["on_var"]};font-size:13px">— Select a project to see its expenditure breakup —</div>'}
    </div>
  </div>

  <!-- Export bar -->
  <div style="display:flex;justify-content:flex-end;margin-bottom:16px">
    <a href="/export_excel/expenditures"
       style="display:inline-flex;align-items:center;gap:6px;padding:7px 14px;
              background:#fff;border:1px solid {C['outline']};border-radius:6px;
              font-size:12px;font-weight:600;color:#374151;text-decoration:none;
              transition:background .15s"
       onmouseover="this.style.background='{C['surface_high']}'"
       onmouseout="this.style.background='#fff'">
      ⬇ Export Excel
    </a>
  </div>
  <!-- ERP Cost Reconciliation -->
  <div class="erp-card" style="margin-bottom:28px">
    <div class="erp-hdr">
      <div style="display:flex;align-items:center;gap:10px">
        <span style="font-size:22px">⚖️</span>
        <span style="font-size:18px;font-weight:700;color:{C['dark']}">ERP Cost Reconciliation</span>
      </div>
      <div style="display:flex;gap:32px;align-items:flex-start">
        <div>
          <div style="font-size:9px;font-weight:700;text-transform:uppercase;
                      letter-spacing:.08em;color:{C['on_var']};margin-bottom:4px">
            Data Available as of
          </div>
          <div style="font-size:13px;font-weight:700;color:{C['dark']}">{data_as_of}</div>
        </div>
        <div>
          <div style="font-size:9px;font-weight:700;text-transform:uppercase;
                      letter-spacing:.08em;color:{C['on_var']};margin-bottom:4px">
            Last ERP Update Date
          </div>
          <div style="font-size:13px;font-weight:700;color:{C['dark']}">{erp_update}</div>
        </div>
      </div>
    </div>

    <form method="post" action="/save_expenditure">
      <input type="hidden" name="co_no"           value="{sel_co}">
      <input type="hidden" name="lob"             value="{sel_lob}">
      <input type="hidden" name="data_as_of"      value="{data_as_of}">
      <input type="hidden" name="erp_update_date" value="{erp_update}">

      <div style="padding:24px 28px">
        <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:20px">
          {erp_cols_html}
        </div>
      </div>

      <!-- Total footer bar -->
      <div style="background:{C['dark']};color:#fff;padding:20px 28px;display:flex;
                  align-items:center;justify-content:space-between;gap:24px;flex-wrap:wrap">
        <div>
          <div style="font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;
                      opacity:.75;margin-bottom:3px">Calculated Total Cumulative Expenditure (A/c)</div>
          <div style="font-size:11px;opacity:.5;font-style:italic">
            Aggregated value for system-wide verification &amp; auditing
          </div>
        </div>
        <div style="display:flex;align-items:center;gap:16px">
          <div style="position:relative;width:210px">
            <span style="position:absolute;left:14px;top:50%;transform:translateY(-50%);
                         font-size:18px;font-weight:700;color:{C['dark']};pointer-events:none">₹</span>
            <input type="text" id="erp-total" readonly value="0.00"
                   style="width:100%;padding:11px 11px 11px 34px;background:#fff;
                          color:{C['dark']};border:none;border-radius:7px;font-size:18px;
                          font-weight:800;outline:none;letter-spacing:-.02em">
          </div>
          <button type="submit"
                  style="background:rgba(255,255,255,.15);color:#fff;
                         border:1.5px solid rgba(255,255,255,.4);border-radius:7px;
                         padding:11px 28px;font-size:11px;font-weight:700;
                         letter-spacing:.07em;text-transform:uppercase;cursor:pointer;
                         backdrop-filter:blur(4px)"
                  onmouseover="this.style.background='rgba(255,255,255,.25)'"
                  onmouseout="this.style.background='rgba(255,255,255,.15)'">
            UPDATE DATA
          </button>
        </div>
      </div>
    </form>
  </div>

  <!-- Add Entry side panel -->
  <div class="side-panel" id="side-panel">
    <div class="sp-hdr">
      <span style="font-size:16px;font-weight:700;color:{C['dark']}">Add Expense Entry</span>
      <button onclick="closePanel()" type="button"
              style="background:none;border:none;font-size:20px;cursor:pointer;
                     color:{C['on_var']};line-height:1">✕</button>
    </div>
    <div class="sp-body">
      <form method="post" action="/save_entry" id="entry-form">
        <input type="hidden" name="co_no"          value="{sel_co}">
        <input type="hidden" name="expenditure_id" value="0">
        <div class="sp-fld">
          <label>Expense Category</label>
          <select name="category" id="sp-cat">{cat_opts}</select>
        </div>
        <div class="sp-fld">
          <label>Description</label>
          <input type="text" name="description" id="sp-desc"
                 placeholder="e.g. Site Infrastructure Supplies">
        </div>
        <div class="sp-g2">
          <div class="sp-fld">
            <label>Quantity</label>
            <input type="number" name="quantity" id="sp-qty"
                   placeholder="0" step="1" min="0" oninput="calcTotal()">
          </div>
          <div class="sp-fld">
            <label>Unit Cost (₹)</label>
            <input type="number" name="unit_cost" id="sp-ucost"
                   placeholder="0.00" step="0.01" min="0" oninput="calcTotal()">
          </div>
        </div>
        <div style="background:#f0f4ff;border:1px solid #c8d8f8;border-radius:8px;
                    padding:16px 18px;margin-top:8px">
          <div style="font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
                      color:{C['on_var']};margin-bottom:6px">Line Total</div>
          <div style="font-size:24px;font-weight:800;color:{C['dark']};
                      letter-spacing:-.02em" id="sp-total">₹ 0.00</div>
        </div>
      </form>
    </div>
    <div class="sp-foot">
      <button onclick="closePanel()" type="button"
              style="flex:1;padding:11px;border:1px solid {C['outline']};border-radius:6px;
                     background:#fff;color:{C['on_var']};font-size:11px;font-weight:700;
                     text-transform:uppercase;letter-spacing:.05em;cursor:pointer">Cancel</button>
      <button onclick="document.getElementById('entry-form').submit()" type="button"
              style="flex:1;padding:11px;background:{C['dark']};color:#fff;border:none;
                     border-radius:6px;font-size:11px;font-weight:700;text-transform:uppercase;
                     letter-spacing:.05em;cursor:pointer">Save Entry</button>
    </div>
  </div>
  <div class="overlay" id="overlay" onclick="closePanel()"></div>

  <!-- Expenditure breakup — head detail modal -->
  <div class="overlay" id="head-overlay" onclick="closeHeadDetail()"></div>
  <div id="head-detail-modal"
       style="position:fixed;top:50%;left:50%;
              transform:translate(-50%,-50%) scale(.96);opacity:0;
              width:680px;max-width:92vw;max-height:80vh;background:#fff;border-radius:12px;
              box-shadow:0 20px 60px rgba(0,0,0,.28);z-index:301;display:none;
              flex-direction:column;transition:opacity .15s ease,transform .15s ease">
    <div style="padding:20px 24px;border-bottom:1px solid {C['outline']};
                display:flex;align-items:center;justify-content:space-between;
                background:{C['surface_low']};border-radius:12px 12px 0 0">
      <div>
        <div style="font-size:9px;font-weight:700;text-transform:uppercase;
                    letter-spacing:.08em;color:{C['on_var']};margin-bottom:3px">Expenditure Head</div>
        <div id="head-modal-title" style="font-size:16px;font-weight:800;color:{C['dark']}">—</div>
      </div>
      <button onclick="closeHeadDetail()" type="button"
              style="background:none;border:none;font-size:20px;cursor:pointer;
                     color:{C['on_var']};line-height:1">✕</button>
    </div>
    <div id="head-modal-body" style="padding:20px 24px;overflow-y:auto;flex:1"></div>
  </div>

  <!-- Hidden form used to delete a line-item entry from within the head-detail modal -->
  <form id="del-entry-form" method="post" action="/delete_entry" style="display:none">
    <input type="hidden" name="entry_id" id="del-entry-id" value="">
    <input type="hidden" name="co_no" value="{sel_co}">
    <input type="hidden" name="lob" value="{sel_lob}">
  </form>

{js}
<script>
(function(){{
  var PAGE = 'expenditures';
  var saved = MonarchFilters.init(PAGE);
  var url   = new URL(window.location.href);
  if (saved && !url.searchParams.get('project_id') && (saved.lob || saved.project_id)) {{
    var qs = [];
    if (saved.lob)        qs.push('lob='        + encodeURIComponent(saved.lob));
    if (saved.project_id) qs.push('project_id=' + encodeURIComponent(saved.project_id));
    if (qs.length) {{ window.location.replace('/expenditures?' + qs.join('&')); return; }}
  }}
  var form = document.getElementById('exp-filter-form');
  if (form) {{
    form.addEventListener('submit', function(){{
      var fd = new FormData(form);
      MonarchFilters.save(PAGE, {{ lob: fd.get('lob')||'', project_id: fd.get('project_id')||'' }});
    }});
  }}
}})();
</script>
</body></html>"""
    return html
# ═══════════════════════════════════════════════════════════════
#  PAGE 4 — FINANCIALS  (UPDATED with child budget/expense rows)
# ═══════════════════════════════════════════════════════════════
def export_financials_pdf_bytes(lob_filter: str = "", ebf_mult: float = 1.35, pm_filter: str = "") -> bytes:
    # A "pm" login only ever sees their own projects — ignore any other
    # pm_filter value that may have been passed (e.g. a hand-edited URL).
    role   = (getattr(_req_ctx, "role", "") or "").lower()
    own_pm = (getattr(_req_ctx, "pm_name", "") or "").strip()
    if role == "pm" and own_pm:
        pm_filter = own_pm
    try:
        ebf_mult = float(ebf_mult)
    except (TypeError, ValueError):
        ebf_mult = 1.35
    projects = fetch_all_projects(lob_filter=lob_filter)
    if pm_filter:
        projects = [p for p in projects if (p.get("project_manager") or "") == pm_filter]
    parents  = [p for p in projects if p.get("project_type","parent") != "child"]
    rows, total_ov = [], 0.0
    for p in parents:
        ov  = _rev_ov(p); itd = _n(p.get("itd_cost")); col = _n(p.get("collection_till_date"))
        ebf = round(ebf_mult*itd, 2)
        total_ov += ov
        rows.append([p.get("co_no","—"), p.get("project_name","—"), p.get("project_manager","—"),
                     f"{ov:,.0f}", f"{itd:,.0f}", f"{col:,.0f}", f"{ebf:,.0f}"])
    return _build_table_pdf(
        "Financial Performance Report",
        f"Project performance analysis — {lob_filter or 'All LOB'}",
        ["CO.No","Project","PM","Order Value","Total Expenditure","Collection",f"EBF ({ebf_mult:g}x)"],
        rows, footer_label="Total Order Value", footer_value=total_ov)


def export_financials_excel_bytes(lob_filter: str = "", ebf_mult: float = 1.35, pm_filter: str = "") -> bytes:
    """Export Financial Performance Report as .xlsx"""
    from openpyxl.styles import Font, PatternFill, Alignment
    # A "pm" login only ever sees their own projects — ignore any other
    # pm_filter value that may have been passed (e.g. a hand-edited URL).
    role   = (getattr(_req_ctx, "role", "") or "").lower()
    own_pm = (getattr(_req_ctx, "pm_name", "") or "").strip()
    if role == "pm" and own_pm:
        pm_filter = own_pm
    try:
        ebf_mult = float(ebf_mult)
    except (TypeError, ValueError):
        ebf_mult = 1.35
    projects = fetch_all_projects(lob_filter=lob_filter)
    if pm_filter:
        projects = [p for p in projects if (p.get("project_manager") or "") == pm_filter]
    parents  = [p for p in projects if p.get("project_type", "parent") != "child"]
    wb = Workbook(); ws = wb.active; ws.title = "Financials"
    hdr_fill = PatternFill("solid", fgColor="000E24")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    alt_fill = PatternFill("solid", fgColor="F2F4F6")
    headers  = ["CO.No", "Project Name", "Project Manager", "Order Value (Rs)",
                "Total Expenditure (Rs)", "Collection (Rs)", f"EBF {ebf_mult:g}x (Rs)", "Completion %"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = hdr_fill; cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    total_ov = 0.0
    for i, p in enumerate(parents, start=2):
        ov  = _rev_ov(p); itd = _n(p.get("itd_cost")); col = _n(p.get("collection_till_date"))
        ebf = round(ebf_mult * itd, 2); total_ov += ov
        ws.append([p.get("co_no",""), p.get("project_name",""), p.get("project_manager",""),
                   ov, itd, col, ebf, _n(p.get("physical_completion"))])
        if i % 2 == 0:
            for cell in ws[i]: cell.fill = alt_fill
    ws.append(["TOTAL", "", "", total_ov, "", "", "", ""])
    for cell in ws[ws.max_row]: cell.font = Font(bold=True, size=10)
    for idx, w in enumerate([10, 40, 22, 16, 16, 16, 16, 12], start=1):
        ws.column_dimensions[ws.cell(1, idx).column_letter].width = w
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf.read()


def export_variations_excel_bytes() -> bytes:
    """Export Parent Project Variations as .xlsx"""
    from openpyxl.styles import Font, PatternFill, Alignment
    db_vars = fetch_all_variations()
    wb = Workbook(); ws = wb.active; ws.title = "Variations"
    hdr_fill = PatternFill("solid", fgColor="000E24")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    alt_fill = PatternFill("solid", fgColor="F2F4F6")
    headers  = ["Child CO.No", "Child Project", "Parent Project",
                "Original Value (Rs)", "Variation Amt (Rs)", "Revised Value (Rs)", "Date", "Status"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = hdr_fill; cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    total_var = 0.0
    for i, v in enumerate(db_vars, start=2):
        amt = _n(v.get("variation_amt")); total_var += amt
        ws.append([v.get("child_co_no",""), v.get("child_name",""), v.get("parent_name",""),
                   _n(v.get("original_value")), amt, _n(v.get("revised_value")),
                   v.get("variation_date",""), v.get("status","—")])
        if i % 2 == 0:
            for cell in ws[i]: cell.fill = alt_fill
    ws.append(["NET VARIATION", "", "", "", total_var, "", "", ""])
    for cell in ws[ws.max_row]: cell.font = Font(bold=True, size=10)
    for idx, w in enumerate([12, 36, 36, 16, 16, 16, 14, 14], start=1):
        ws.column_dimensions[ws.cell(1, idx).column_letter].width = w
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf.read()


def export_kpi_excel_bytes() -> bytes:
    """Export KPI Red Flag Monitor as .xlsx"""
    from openpyxl.styles import Font, PatternFill, Alignment
    # A "pm" login only ever sees their own projects.
    role   = (getattr(_req_ctx, "role", "") or "").lower()
    own_pm = (getattr(_req_ctx, "pm_name", "") or "").strip()
    projects = fetch_all_projects()
    if role == "pm" and own_pm:
        projects = [pr for pr in projects if (pr.get("project_manager") or "") == own_pm]
    wb = Workbook(); ws = wb.active; ws.title = "KPI"
    hdr_fill = PatternFill("solid", fgColor="000E24")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    alt_fill = PatternFill("solid", fgColor="F2F4F6")
    headers  = ["CO.No", "Project Name", "LOB", "Status",
                "Order Value (Rs)", "Total Expenditure (Rs)", "Collection (Rs)",
                "EAC (Rs)", "Physical %", "Billed Amt (Rs)"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = hdr_fill; cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for i, p in enumerate(projects, start=2):
        ws.append([
            p.get("co_no",""), p.get("project_name",""),
            p.get("hier_business_line") or p.get("business_line",""),
            p.get("project_status",""),
            _rev_ov(p), _n(p.get("itd_cost")), _n(p.get("collection_till_date")),
            _n(p.get("rev_latest_eac") or p.get("baseline_eac")),
            _n(p.get("physical_completion")), _n(p.get("billed_amount")),
        ])
        if i % 2 == 0:
            for cell in ws[i]: cell.fill = alt_fill
    for idx, w in enumerate([10, 40, 20, 14, 16, 16, 16, 16, 10, 16], start=1):
        ws.column_dimensions[ws.cell(1, idx).column_letter].width = w
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf.read()


def export_billing_excel_bytes(project_id=None, lob=None, pending_only=False) -> bytes:
    """Export Billing & Collection entries as .xlsx"""
    from openpyxl.styles import Font, PatternFill, Alignment
    entries = fetch_billing_entries_filtered(
        project_id=project_id or None,
        lob=lob or None,
        pending_only=pending_only
    )
    wb = Workbook(); ws = wb.active; ws.title = "Billing & Collection"
    hdr_fill = PatternFill("solid", fgColor="000E24")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    alt_fill = PatternFill("solid", fgColor="F2F4F6")
    headers  = ["RA No.", "Project ID", "LOB", "Billing Date",
                "Billed Amount (Rs)", "Collection Date", "Collection Amount (Rs)"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = hdr_fill; cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    total_billed = 0.0; total_collected = 0.0
    for i, e in enumerate(entries, start=2):
        billed    = _n(e.get("billed_amt"))
        collected = _n(e.get("collection_amt"))
        total_billed += billed; total_collected += collected
        ws.append([e.get("ra_no",""), e.get("project_id",""), e.get("lob",""),
                   e.get("billing_date",""), billed,
                   e.get("collection_date",""), collected])
        if i % 2 == 0:
            for cell in ws[i]: cell.fill = alt_fill
    ws.append(["TOTAL", "", "", "", total_billed, "", total_collected])
    for cell in ws[ws.max_row]: cell.font = Font(bold=True, size=10)
    for idx, w in enumerate([14, 12, 20, 14, 18, 14, 20], start=1):
        ws.column_dimensions[ws.cell(1, idx).column_letter].width = w
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf.read()

def financials(alert_html="", prefill=None, username="Admin") -> str:
    p       = prefill or {}
    role    = (getattr(_req_ctx, "role", "") or "").lower()
    is_pm_role = role == "pm"
    own_pm  = (getattr(_req_ctx, "pm_name", "") or "").strip()
    is_scoped_pm = bool(is_pm_role and own_pm)

    lob_all = ["All Lines of Business"] + get_visible_lob_list()
    sel_lob = p.get("lob", "All Lines of Business")
    lob_opts = "".join(
        f'<option {"selected" if lo == sel_lob else ""}>{lo}</option>'
        for lo in lob_all
    )

    # A "pm" login only ever sees their own projects — ignore any other
    # pm= value that may have been passed (e.g. a hand-edited URL/form),
    # and don't offer the other-PM options in the selector at all.
    if is_scoped_pm:
        sel_pm = own_pm
        pm_selector_html = f"""
        <div style="background:{C['surface_low']};color:{C['dark']};font-size:14px;font-weight:600;
                    padding:0 12px;border-radius:6px;display:inline-flex;align-items:center;gap:6px">
          👤 {_htmlmod.escape(own_pm)}
        </div>
        <input type="hidden" name="pm" value="{_htmlmod.escape(own_pm)}">"""
    else:
        pm_all  = ["All Project Managers"] + get_project_managers()
        sel_pm  = p.get("pm", "All Project Managers")
        pm_opts = "".join(
            f'<option {"selected" if pmn == sel_pm else ""}>{pmn}</option>'
            for pmn in pm_all
        )
        pm_selector_html = f"""
      <select name="pm" onchange="this.form.submit()"
              style="border:none;background:transparent;font-size:14px;color:{C['dark']};
                     outline:none;cursor:pointer;min-width:200px;font-weight:500">
        {pm_opts}
      </select>"""

    lob_filt = "" if sel_lob == "All Lines of Business" else sel_lob
    projects  = fetch_all_projects(lob_filter=lob_filt)
    if sel_pm != "All Project Managers":
        projects = [pr for pr in projects if (pr.get("project_manager") or "") == sel_pm]

    # ── EBF multiplier: use the value saved for this LOB (Financials → Set),
    # falling back to the platform default (1.35) when nothing's been set. ──
    active_ebf_mult = get_ebf_mult(sel_lob)

    # ── Only parent projects in the main table ──
    parent_list = [pr for pr in projects if pr.get("project_type", "parent") != "child"]
    all_projects_count = len(fetch_all_projects())  # total unfiltered for "Showing X of Y"

    # ── Sub-projects (children), same linkage rules used by Consolidated Reports ──
    _fin_all_children = _rep_child_projects(fetch_all_projects())

    # ── Totals ──
    total_ov  = sum(_n(pr.get("revised_order_value") or (_n(pr.get("order_value")) + _n(pr.get("variation_amt")))) for pr in parent_list)
    total_itd = sum(_n(pr.get("itd_cost"))             for pr in parent_list)   # Col Z
    total_col = sum(_n(pr.get("collection_till_date")) for pr in parent_list)
    total_ebf = round(active_ebf_mult * total_itd, 2)                           # Col AA (LOB-specific EBF)

    # ── FIXED: avg_perf = Collection as % of Order Value (display KPI) ──
    avg_perf = (total_col / total_ov * 100) if total_ov else 0

    # ── FIXED: ratio = Collection / EBF  (meaningful: ≥1.0 means collection covers EBF) ──
    # 1.35 is the EBF multiplier; if Collection >= EBF it means we've recovered costs+margin
    ratio     = total_col / total_ebf if total_ebf else 0
    ratio_pct = ratio * 100  # for display purposes

    # ── Threshold: ratio >= 1.0 means collection has covered the 1.35x cost target ──
    ratio_str     = f"{ratio:.2f}x"
    ratio_lbl     = "Profit — Collection ≥ EBF" if ratio >= 1.0 else "Loss — Collection < EBF"
    ratio_badge_cls = "background:#f0fdf4;color:#166534" if ratio >= 1.0 else "background:#fef2f2;color:#991b1b"

    # Sub-label showing how far above/below threshold
    if ratio >= 1.0:
        surplus = total_col - total_ebf
        ratio_sub = f"Surplus of ₹{surplus/1e5:,.1f}L above EBF target"
    else:
        deficit = total_ebf - total_col
        ratio_sub = f"Deficit of ₹{deficit/1e5:,.1f}L below EBF target"

    # ── Vertical gauge bar, scale 1.00x → active EBF multiplier in 0.05 steps ──
    GAUGE_MIN, GAUGE_MAX, GAUGE_STEP = 1.00, round(max(1.35, active_ebf_mult), 2), 0.05
    gauge_ticks = []
    t = GAUGE_MIN
    while t <= GAUGE_MAX + 1e-9:
        gauge_ticks.append(round(t, 2))
        t += GAUGE_STEP

    def gauge_pos(val):
        """% from bottom (0%) to top (100%) for a value on the GAUGE_MIN..GAUGE_MAX scale, clamped."""
        v = max(GAUGE_MIN, min(GAUGE_MAX, val))
        return (v - GAUGE_MIN) / (GAUGE_MAX - GAUGE_MIN) * 100

    ratio_clamped      = max(GAUGE_MIN, min(GAUGE_MAX, ratio))
    marker_pos_pct     = gauge_pos(ratio)
    marker_color       = "#166534" if ratio >= 1.0 else "#dc2626"
    out_of_range_note  = ""
    if ratio < GAUGE_MIN:
        out_of_range_note = f'<div id="gauge-oor" style="font-size:9px;color:#991b1b;margin-top:4px">Below {GAUGE_MIN:.2f}x — pinned to scale floor</div>'
    elif ratio > GAUGE_MAX:
        out_of_range_note = f'<div id="gauge-oor" style="font-size:9px;color:#166534;margin-top:4px">Above {GAUGE_MAX:.2f}x — pinned to scale ceiling</div>'
    else:
        out_of_range_note = '<div id="gauge-oor" style="font-size:9px;margin-top:4px"></div>'

    # Tick marks (rendered bottom→top, so reverse the list for top-down DOM order is not needed
    # since we position with `bottom: %` absolutely)
    gauge_tick_html = ""
    for gt in gauge_ticks:
        pos = gauge_pos(gt)
        is_break_even = abs(gt - 1.0) < 1e-9
        tick_color = "#166534" if is_break_even else "#9ca3af"
        tick_weight = "800" if is_break_even else "600"
        gauge_tick_html += f"""
          <div style="position:absolute;bottom:{pos:.2f}%;left:0;right:0;display:flex;align-items:center;
                      gap:6px;transform:translateY(50%)">
            <div style="width:9px;height:{2 if is_break_even else 1}px;background:{tick_color}"></div>
            <span style="font-size:9px;font-weight:{tick_weight};color:{tick_color};
                         font-family:'Courier New',monospace">{gt:.2f}x</span>
          </div>"""

    gauge_html = f"""
        <div style="text-align:center;min-width:170px">
          <div style="font-size:11px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;
                      color:{C['on_var']};margin-bottom:10px">Collection / EBF</div>

          <div style="display:flex;justify-content:center;align-items:flex-end;gap:14px">
            <!-- Vertical scale track -->
            <div style="position:relative;width:16px;height:200px;border-radius:99px;
                        background:linear-gradient(to top, #dc2626 0%, #f59e0b 45%, #166534 100%);
                        box-shadow:inset 0 0 0 1px {C['outline']};overflow:visible">
              <!-- Marker (current ratio position) -->
              <div id="gauge-marker" title="Current: {ratio_str}"
                   style="position:absolute;left:50%;bottom:{marker_pos_pct:.2f}%;
                          transform:translate(-50%,50%);width:26px;height:26px;
                          border-radius:50%;background:#fff;border:3px solid {marker_color};
                          box-shadow:0 2px 6px rgba(0,0,0,.25);z-index:2;
                          display:flex;align-items:center;justify-content:center">
                <div id="gauge-marker-dot" style="width:6px;height:6px;border-radius:50%;background:{marker_color}"></div>
              </div>
            </div>

            <!-- Tick labels -->
            <div style="position:relative;width:48px;height:200px">
              {gauge_tick_html}
            </div>
          </div>

          <div id="gauge-ratio-str" style="font-size:30px;font-weight:800;color:{marker_color};line-height:1;margin-top:14px">{ratio_str}</div>
          <span id="gauge-ratio-badge" style="display:inline-block;padding:5px 18px;border-radius:6px;font-size:12px;
                       font-weight:700;margin-top:8px;{ratio_badge_cls}">{ratio_lbl}</span>
          <div id="gauge-avg-perf" style="font-size:10px;color:#9ca3af;margin-top:6px">
            Avg Collection Rate: {avg_perf:.1f}% of Order Value
          </div>
          {out_of_range_note}
        </div>"""

    n_rows = len(parent_list)
    INITIAL_ROWS_SHOWN = 10

    def to_lakhs(v):
        """Convert raw ₹ value to Lakhs, formatted to 2 decimal places."""
        return round(_n(v) / 1e5, 2)

    # ── Build table rows ──
    proj_rows_html = ""
    for _row_idx, pr in enumerate(parent_list):
        co   = pr.get("co_no")           or "—"
        name = pr.get("project_name")    or "—"
        pm   = pr.get("project_manager") or "—"
        st   = pr.get("project_status")  or "Active"
        pid  = pr.get("id", "")

        rev_ov = _n(pr.get("revised_order_value")) or (_n(pr.get("order_value")) + _n(pr.get("variation_amt")))
        ov_l  = to_lakhs(rev_ov)                          # Col I — Revised Order Value
        itd_l = to_lakhs(pr.get("itd_cost"))              # Col Z — Total ITD Cost
        col_l = to_lakhs(pr.get("collection_till_date"))  # Col S
        ebf_l = round(active_ebf_mult * itd_l, 2)         # Col AA = LOB EBF multiplier × ITD

          # ── Consolidated-Report-style fields ──
        billed_l = to_lakhs(pr.get("billed_amount"))
        bnc_l    = to_lakhs(max(_n(pr.get("billed_amount")) - _n(pr.get("collection_till_date")), 0))
        ub_l     = to_lakhs(max(_rep_oi_amt(pr) - _n(pr.get("billed_amount")), 0))
        fin_pct  = _rep_fin_pct(pr)
        phy_pct  = _rep_phy_pct(pr)
        dur_days = _rep_duration_days(pr)
        if dur_days is None:
            dur_str, dur_style = "—", f"color:{C['on_var']}"
        elif dur_days < 0:
            dur_str, dur_style = f"Overdue {abs(dur_days)}d", f"color:{C['error']};font-weight:700"
        else:
            dur_str, dur_style = f"{dur_days}d", f"color:{C['on_surface']}"

        ebf_l = round(active_ebf_mult * itd_l, 2)         # Col AA = LOB EBF multiplier × ITD

      

        # ── Collection Performance = (Collection - EBF) / EBF × 100
        # Positive = collection exceeded the billing floor (PROFIT)
        # Negative = collection is short of the billing floor (LOSS)
        if ebf_l:
            perf = (col_l - ebf_l) / ebf_l * 100
        else:
            perf = 0.0

        # ── EBF badge color based on col_l vs ebf_l ──
        if col_l >= ebf_l:
            ebf_cls = "color:#166534;background:#f0fdf4;border:1px solid #bbf7d0"
        elif col_l >= ebf_l * 0.80:
            ebf_cls = "color:#92400e;background:#fffbeb;border:1px solid #fde68a"
        else:
            ebf_cls = "color:#991b1b;background:#fef2f2;border:1px solid #fecaca"

        # Coll. Perf sign + color
        if perf >= 0:
            perf_str   = f"+{perf:.1f}%"
            perf_color = "#166534"
            perf_bg    = "#f0fdf4"
        else:
            perf_str   = f"{perf:.1f}%"
            perf_color = "#991b1b"
            perf_bg    = "#fef2f2"

        # Status badge
        if st in ("Cancelled", "Hold"):
            st_cls = "background:#ffdad6;color:#ba1a1a"
        elif st in ("Closed", "Execution Closed"):
            st_cls = f"background:{C['surface_high']};color:{C['on_var']}"
        else:
            st_cls = "background:#fde68a;color:#92400e"

        fp = (pr.get('folder_path') or '').replace("\\", "\\\\").replace("'", "\\'")
        folder_link = (
            f'&nbsp;<a href="javascript:void(0)" onclick="openProjectFolder(\'{fp}\')" '
            f'style="color:{C["secondary"]};font-size:12px;font-weight:700;text-decoration:none" '
            f'title="{fp}">📂 Folder</a>'
        ) if fp else ""

                # data-itd / data-col / data-ov / data-billed / data-bnc / data-ub / data-finpct /
        # data-phypct / data-dur carry the raw per-row figures so both the EBF multiplier
        # slider and the name-search filter can recompute totals client-side without a reload.
        _row_hidden_style = "" if _row_idx < INITIAL_ROWS_SHOWN else ' style="display:none"'
        _search_blob = f"{name} {co}".lower().replace('"', '').replace("'", "")
        _dur_attr = "" if dur_days is None else dur_days
        proj_rows_html += f"""
          <tr class="fin-proj-row" data-row-idx="{_row_idx}" data-itd="{itd_l}" data-col="{col_l}" data-ov="{ov_l}" data-billed="{billed_l}" data-bnc="{bnc_l}" data-ub="{ub_l}" data-finpct="{fin_pct}" data-phypct="{phy_pct}" data-dur="{_dur_attr}" data-search="{_search_blob}"{_row_hidden_style}>
            <td class="fin-freeze-1" style="padding:12px 14px;font-weight:600;color:{C['dark']};font-size:12px;min-width:90px">{co}</td>
            <td class="fin-freeze-2" style="padding:12px 14px;font-weight:600;color:{C['dark']};font-size:12px;min-width:200px">{name}</td>
            <td data-parentcol style="display:none;padding:12px 14px;font-size:12px;color:#6b7280;width:230px;max-width:230px;white-space:normal;line-height:1.4">—</td>
            <td style="padding:12px 14px;font-size:12px;color:{C['on_var']}">{pm}</td>
            <td data-col="ov" style="padding:12px 14px;text-align:right;font-family:'Courier New',monospace;font-size:12px">{ov_l:,.2f}</td>
            <td data-col="exp" style="padding:12px 14px;text-align:right;font-family:'Courier New',monospace;font-size:12px">{itd_l:,.2f}</td>
            <td data-col="billed" style="padding:12px 14px;text-align:right;font-family:'Courier New',monospace;font-size:12px">{billed_l:,.2f}</td>
            <td data-col="col" style="padding:12px 14px;text-align:right;font-family:'Courier New',monospace;font-size:12px">{col_l:,.2f}</td>
            <td data-col="bnc" style="padding:12px 14px;text-align:right;font-family:'Courier New',monospace;font-size:12px">{bnc_l:,.2f}</td>
            <td data-col="ub" style="padding:12px 14px;text-align:right;font-family:'Courier New',monospace;font-size:12px">{ub_l:,.2f}</td>
            <td data-col="finpct" style="padding:12px 14px;text-align:center;font-size:12px">{fin_pct:.1f}%</td>
            <td data-col="phypct" style="padding:12px 14px;text-align:center;font-size:12px">{phy_pct:.1f}%</td>
            <td data-col="dur" style="padding:12px 14px;text-align:center;font-size:12px;{dur_style}">{dur_str}</td>
            <td data-col="ebf" style="padding:12px 14px;text-align:center">
              <span class="ebf-badge" style="display:inline-block;padding:3px 10px;border-radius:4px;font-size:11px;font-weight:700;{ebf_cls}">{ebf_l:,.2f}</span>
            </td>
            <td data-col="collperf" style="padding:12px 14px;text-align:center">
              <span class="perf-badge" style="font-weight:700;color:{perf_color};background:{perf_bg};padding:3px 8px;border-radius:4px;font-size:11px">{perf_str}</span>
            </td>
          </tr>"""

        # ── Sub-project (child) rows for this parent — hidden by default, shown
        # only in "Sub Projects" view. Same layout/columns as the parent row so
        # the existing column-toggle chips work identically on them. ──
        _kids = _rep_kids_of(pr, _fin_all_children)
        for _kid_idx, kid in enumerate(_kids):
            k_co   = kid.get("co_no")           or "—"
            k_name = kid.get("project_name")    or "—"
            k_pm   = kid.get("project_manager") or "—"

            k_rev_ov = _n(kid.get("revised_order_value")) or (_n(kid.get("order_value")) + _n(kid.get("variation_amt")))
            k_ov_l   = to_lakhs(k_rev_ov)
            k_itd_l  = to_lakhs(kid.get("itd_cost"))
            k_col_l  = to_lakhs(kid.get("collection_till_date"))
            k_ebf_l  = round(active_ebf_mult * k_itd_l, 2)

            k_billed_l = to_lakhs(kid.get("billed_amount"))
            k_bnc_l    = to_lakhs(max(_n(kid.get("billed_amount")) - _n(kid.get("collection_till_date")), 0))
            k_ub_l     = to_lakhs(max(_rep_oi_amt(kid) - _n(kid.get("billed_amount")), 0))
            k_fin_pct  = _rep_fin_pct(kid)
            k_phy_pct  = _rep_phy_pct(kid)
            k_dur_days = _rep_duration_days(kid)
            if k_dur_days is None:
                k_dur_str, k_dur_style = "—", f"color:{C['on_var']}"
            elif k_dur_days < 0:
                k_dur_str, k_dur_style = f"Overdue {abs(k_dur_days)}d", f"color:{C['error']};font-weight:700"
            else:
                k_dur_str, k_dur_style = f"{k_dur_days}d", f"color:{C['on_surface']}"

            if k_ebf_l:
                k_perf = (k_col_l - k_ebf_l) / k_ebf_l * 100
            else:
                k_perf = 0.0

            if k_col_l >= k_ebf_l:
                k_ebf_cls = "color:#166534;background:#f0fdf4;border:1px solid #bbf7d0"
            elif k_col_l >= k_ebf_l * 0.80:
                k_ebf_cls = "color:#92400e;background:#fffbeb;border:1px solid #fde68a"
            else:
                k_ebf_cls = "color:#991b1b;background:#fef2f2;border:1px solid #fecaca"

            if k_perf >= 0:
                k_perf_str, k_perf_color, k_perf_bg = f"+{k_perf:.1f}%", "#166534", "#f0fdf4"
            else:
                k_perf_str, k_perf_color, k_perf_bg = f"{k_perf:.1f}%", "#991b1b", "#fef2f2"

            k_fp = (kid.get('folder_path') or '').replace("\\", "\\\\").replace("'", "\\'")
            k_folder_link = (
                f'&nbsp;<a href="javascript:void(0)" onclick="openProjectFolder(\'{k_fp}\')" '
                f'style="color:{C["secondary"]};font-size:12px;font-weight:700;text-decoration:none" '
                f'title="{k_fp}">📂 Folder</a>'
            ) if k_fp else ""

            _parent_label = f"{co} — {name}" if co and co != "—" else name
            _k_search_blob = f"{k_name} {k_co} {name} {co}".lower().replace('"', '').replace("'", "")
            _k_dur_attr = "" if k_dur_days is None else k_dur_days
            proj_rows_html += f"""
          <tr class="fin-child-row" data-itd="{k_itd_l}" data-col="{k_col_l}" data-ov="{k_ov_l}" data-billed="{k_billed_l}" data-bnc="{k_bnc_l}" data-ub="{k_ub_l}" data-finpct="{k_fin_pct}" data-phypct="{k_phy_pct}" data-dur="{_k_dur_attr}" data-search="{_k_search_blob}" style="display:none">
            <td class="fin-freeze-1" style="padding:12px 14px;font-weight:600;color:{C['dark']};font-size:12px;min-width:90px">{k_co}</td>
            <td class="fin-freeze-2" style="padding:12px 14px;font-weight:600;color:{C['dark']};font-size:12px;min-width:200px">{k_name}</td>
            <td data-parentcol style="display:none;padding:12px 14px;font-size:12px;color:#6b7280;width:230px;max-width:230px;white-space:normal;line-height:1.4">{_parent_label}</td>
            <td style="padding:12px 14px;font-size:12px;color:{C['on_var']}">{k_pm}</td>
            <td data-col="ov" style="padding:12px 14px;text-align:right;font-family:'Courier New',monospace;font-size:12px">{k_ov_l:,.2f}</td>
            <td data-col="exp" style="padding:12px 14px;text-align:right;font-family:'Courier New',monospace;font-size:12px">{k_itd_l:,.2f}</td>
            <td data-col="billed" style="padding:12px 14px;text-align:right;font-family:'Courier New',monospace;font-size:12px">{k_billed_l:,.2f}</td>
            <td data-col="col" style="padding:12px 14px;text-align:right;font-family:'Courier New',monospace;font-size:12px">{k_col_l:,.2f}</td>
            <td data-col="bnc" style="padding:12px 14px;text-align:right;font-family:'Courier New',monospace;font-size:12px">{k_bnc_l:,.2f}</td>
            <td data-col="ub" style="padding:12px 14px;text-align:right;font-family:'Courier New',monospace;font-size:12px">{k_ub_l:,.2f}</td>
            <td data-col="finpct" style="padding:12px 14px;text-align:center;font-size:12px">{k_fin_pct:.1f}%</td>
            <td data-col="phypct" style="padding:12px 14px;text-align:center;font-size:12px">{k_phy_pct:.1f}%</td>
            <td data-col="dur" style="padding:12px 14px;text-align:center;font-size:12px;{k_dur_style}">{k_dur_str}</td>
            <td data-col="ebf" style="padding:12px 14px;text-align:center">
              <span class="ebf-badge" style="display:inline-block;padding:3px 10px;border-radius:4px;font-size:11px;font-weight:700;{k_ebf_cls}">{k_ebf_l:,.2f}</span>
            </td>
            <td data-col="collperf" style="padding:12px 14px;text-align:center">
              <span class="perf-badge" style="font-weight:700;color:{k_perf_color};background:{k_perf_bg};padding:3px 8px;border-radius:4px;font-size:11px">{k_perf_str}</span>
            </td>
          </tr>"""

    n_sub_rows = sum(len(_rep_kids_of(pr, _fin_all_children)) for pr in parent_list)

    if not proj_rows_html:
        proj_rows_html = (
            f'<tr><td colspan="15" style="padding:28px;text-align:center;color:#aaa">'
            f'No projects found. <a href="/data_entry" style="color:{C["dark"]};font-weight:700">Add projects →</a>'
            f'</td></tr>'
        )
    # ── Totals in Lakhs ──
    total_ov_l  = round(total_ov  / 1e5, 2)
    total_itd_l = round(total_itd / 1e5, 2)
    total_col_l = round(total_col / 1e5, 2)
    total_ebf_l = round(active_ebf_mult * total_itd_l, 2)

      # ── Consolidated-Report-style totals (Billed / Unbilled / % averages) ──
    total_billed_raw   = sum(_n(pr.get("billed_amount")) for pr in parent_list)
    total_bnc_raw      = sum(max(_n(pr.get("billed_amount")) - _n(pr.get("collection_till_date")), 0) for pr in parent_list)
    total_unbilled_raw = sum(max(_rep_oi_amt(pr) - _n(pr.get("billed_amount")), 0) for pr in parent_list)
    total_billed_l    = round(total_billed_raw   / 1e5, 2)
    total_bnc_l       = round(total_bnc_raw      / 1e5, 2)
    total_unbilled_l  = round(total_unbilled_raw / 1e5, 2)
    avg_fin_pct = (sum(_rep_fin_pct(pr) for pr in parent_list) / len(parent_list)) if parent_list else 0
    avg_phy_pct = (sum(_rep_phy_pct(pr) for pr in parent_list) / len(parent_list)) if parent_list else 0
    _dur_vals   = [d for d in (_rep_duration_days(pr) for pr in parent_list) if d is not None]
    avg_dur     = (sum(_dur_vals) / len(_dur_vals)) if _dur_vals else None
    avg_dur_str = "—" if avg_dur is None else (f"Overdue {abs(round(avg_dur))}d" if avg_dur < 0 else f"{round(avg_dur)}d")

    # Total footer Coll. Perf = (total_col - total_ebf) / total_ebf * 100
    total_perf = (total_col - total_ebf) / total_ebf * 100 if total_ebf else 0
    if total_perf >= 0:
        total_perf_str   = f"+{total_perf:.1f}%"
        total_perf_color = "#166534"
    else:
        total_perf_str   = f"{total_perf:.1f}%"
        total_perf_color = "#991b1b"

    _fin_scope_label_js = (
        f"({sel_lob if sel_lob != 'All Lines of Business' else 'All LOB'})"
    ).replace("\\", "\\\\").replace('"', '\\"')

    css = f"""<style>
.fin-wrap{{background:#fff;border:1px solid {C['outline']};border-radius:12px;overflow:hidden;margin-bottom:20px}}
.fin-thead th{{padding:10px 14px;font-size:10px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;
               color:{C['on_var']};background:{C['surface_low']};border-bottom:2px solid {C['outline']};white-space:nowrap;
               position:sticky;top:0;z-index:2}}
.fin-tfoot td{{padding:12px 14px;font-size:11px;font-weight:700;font-family:'Courier New',monospace;
               background:{C['surface_high']};border-top:2px solid {C['outline']}}}

/* ── Frozen (sticky) leading columns: CO.No + Project Name stay put
   while the rest of the table scrolls horizontally underneath them ── */
.fin-wrap{{position:relative}}
.fin-scroll{{overflow-x:auto;overflow-y:auto;max-height:65vh}}
.fin-freeze-1{{position:sticky;left:0;z-index:3}}
.fin-freeze-2{{position:sticky;left:90px;z-index:3;box-shadow:2px 0 4px -2px rgba(0,0,0,.15)}}
thead .fin-freeze-1,thead .fin-freeze-2{{z-index:4;background:{C['surface_low']}}}
tbody .fin-freeze-1,tbody .fin-freeze-2{{background:#fff}}
tfoot .fin-freeze-1,tfoot .fin-freeze-2{{background:{C['surface_high']};z-index:4}}

/* ── EBF multiplier vertical scroller ── */
.ebf-mult-wrap{{display:flex;flex-direction:column;align-items:center;gap:6px;padding:6px 4px}}
.ebf-mult-wrap input[type=range]{{
  writing-mode: vertical-lr;
  direction: rtl;
  width: 18px;
  height: 90px;
  accent-color: {C['dark']};
  cursor: pointer;
}}
.ebf-mult-label{{font-size:11px;font-weight:800;font-family:'Courier New',monospace;color:{C['dark']}}}

/* ── Column toggle chips (same pattern as Consolidated Reports) ── */
.fin-tabs-bar{{display:flex;align-items:center;gap:12px;flex-wrap:wrap;
    padding:12px 20px;border-bottom:1px solid {C['outline']};background:#fff}}
.fin-tabs-label{{font-size:11px;font-weight:700;letter-spacing:.06em;
    text-transform:uppercase;color:{C['on_var']};white-space:nowrap}}
.fin-tabs{{display:flex;gap:6px;flex-wrap:wrap;align-items:center}}
.fin-tab{{padding:7px 16px;border-radius:20px;border:1px solid {C['outline']};
    font-size:12px;font-weight:600;cursor:pointer;background:{C['surface_low']};
    color:{C['on_var']};transition:all .15s;white-space:nowrap;
    display:inline-flex;align-items:center;gap:6px}}
.fin-tab::before{{content:'';width:7px;height:7px;border-radius:50%;
    background:{C['outline']};transition:background .15s}}
.fin-tab:hover{{background:{C['surface_high']};color:{C['dark']}}}
.fin-tab.active{{background:{C['dark']};color:#fff;border-color:{C['dark']}}}
.fin-tab.active::before{{background:#22c55e}}
.fin-reset-link{{font-size:11px;font-weight:700;color:{C['secondary']};cursor:pointer;
    text-decoration:underline;white-space:nowrap;background:none;border:none;margin-left:auto}}
</style>"""

    html  = _html_head("Financial Performance") + css + "<body>"
    html += _sidebar("financials") + _topbar("FINANCIAL PERFORMANCE", username)
    html += f"""
<div class="main"><div class="canvas" style="max-width:1400px">
  {alert_html}

  <!-- Page header -->
  <div style="display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:24px">
    <div>
      <h2 style="font-size:28px;font-weight:700;color:{C['dark']}">Financial Performance Report</h2>
      <p style="font-size:14px;color:#6b7280;margin-top:4px">Project-level cost, collection and EBF deviation analysis.</p>
    </div>
    <div style="display:flex;gap:10px">
       <a id="export-excel-link" href="/export_excel/financials?lob={sel_lob if sel_lob != 'All Lines of Business' else ''}&pm={sel_pm if sel_pm != 'All Project Managers' else ''}&ebf_mult={active_ebf_mult}"
         style="display:inline-flex;align-items:center;gap:6px;padding:8px 16px;
                background:{C['dark']};color:#fff;border-radius:7px;font-size:12px;
                font-weight:700;text-decoration:none;letter-spacing:.04em"
      >⬇ Export Excel</a>
    </div>
  </div>

  <!-- LOB & PM filter -->
  <form method="post" action="/financials_filter">
    <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;padding:12px 20px;
                display:flex;align-items:center;gap:16px;margin-bottom:24px;flex-wrap:wrap">
      <span style="font-size:18px">≡</span>
      <span style="font-size:14px;font-weight:600;color:{C['dark']}">Line of Business</span>
      <div style="height:20px;width:1px;background:{C['outline']}"></div>
      <select name="lob" onchange="this.form.submit()"
              style="border:none;background:transparent;font-size:14px;color:{C['dark']};
                     outline:none;cursor:pointer;min-width:240px;font-weight:500">
        {lob_opts}
      </select>
      <div style="height:20px;width:1px;background:{C['outline']}"></div>
      <span style="font-size:14px;font-weight:600;color:{C['dark']}">Project Manager</span>
      {pm_selector_html}
    </div>
  </form>

  <!-- Main table card -->
  <div class="fin-wrap">
    <!-- Table header bar -->
    <div style="padding:16px 20px;border-bottom:1px solid {C['outline']};
                display:flex;align-items:center;justify-content:space-between;
                background:{C['surface_low']};flex-wrap:wrap;gap:12px">
      <div style="display:flex;align-items:center;gap:18px;flex-wrap:wrap">
        <span style="font-size:16px;font-weight:700;color:{C['dark']}">Project Performance Analysis</span>
        <form method="post" action="/save_ebf_lob" style="display:flex;align-items:center;gap:8px">
          <input type="hidden" name="lob" value="{sel_lob}">
          <input type="hidden" name="pm" value="{sel_pm}">
          <span style="font-size:11px;font-weight:700;color:{C['on_var']};text-transform:uppercase;letter-spacing:.03em">
            EBF for {"this LOB" if sel_lob != "All Lines of Business" else "All LOBs"}:
          </span>
          <input type="number" name="ebf_mult" id="ebf-mult-input" step="0.01" min="1" value="{active_ebf_mult}"
                 {"disabled" if sel_lob == "All Lines of Business" else ""}
                 style="width:72px;height:32px;border:1px solid {C['outline']};border-radius:6px;
                        padding:0 8px;font-size:12px;text-align:center;font-family:'Courier New',monospace">
          <button type="submit" {"disabled" if sel_lob == "All Lines of Business" else ""}
                  style="height:32px;padding:0 16px;background:{C['dark']};color:#fff;border:none;
                         border-radius:6px;font-size:11px;font-weight:700;letter-spacing:.04em;
                         text-transform:uppercase;cursor:pointer;
                         {'opacity:.45;cursor:not-allowed;' if sel_lob == 'All Lines of Business' else ''}">
            Set
          </button>
        </form>
        {f'<span style="font-size:10.5px;color:#9ca3af">Select a specific LOB above to set its EBF multiplier</span>' if sel_lob == "All Lines of Business" else ''}
      </div>
        <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
        <div style="position:relative">
          <span style="position:absolute;left:10px;top:50%;transform:translateY(-50%);
                       font-size:13px;color:#9ca3af;pointer-events:none">🔍</span>
          <input type="text" id="fin-search" placeholder="Search project name…"
                 oninput="filterFinTable()" autocomplete="off"
                 style="padding:7px 10px 7px 30px;border:1px solid {C['outline']};
                        border-radius:6px;font-size:12px;outline:none;min-width:220px;
                        color:{C['dark']}">
        </div>
        <span id="fin-count-badge" style="font-size:11px;font-weight:700;color:{C['on_var']};
                     background:{C['surface_high']};border:1px solid {C['outline']};
                     padding:4px 12px;border-radius:20px;white-space:nowrap">
          Showing {n_rows} of {all_projects_count} Projects
        </span>
      </div>
    </div>

    <!-- Column toggle chips (add/remove metric columns, like Consolidated Reports) -->
    <div class="fin-tabs-bar">
      <button class="fin-tab" id="fin-subproj-toggle" type="button" onclick="toggleFinSubView()">🌿 Sub Projects</button>
      <div style="width:1px;height:20px;background:{C['outline']}"></div>
      <span class="fin-tabs-label">Columns</span>
      <div class="fin-tabs">
        <button class="fin-tab" data-col="ov"       onclick="toggleFinCol('ov')">Order Value</button>
        <button class="fin-tab" data-col="exp"      onclick="toggleFinCol('exp')">Total Expenditure</button>
        <button class="fin-tab" data-col="billed"   onclick="toggleFinCol('billed')">Billed Amt</button>
        <button class="fin-tab" data-col="col"      onclick="toggleFinCol('col')">Collection</button>
        <button class="fin-tab" data-col="bnc"      onclick="toggleFinCol('bnc')">Billed, Not Collected</button>
        <button class="fin-tab" data-col="ub"       onclick="toggleFinCol('ub')">Unbilled</button>
        <button class="fin-tab" data-col="finpct"   onclick="toggleFinCol('finpct')">% Expenditure Over Order Value</button>
        <button class="fin-tab" data-col="phypct"   onclick="toggleFinCol('phypct')">Physical %</button>
        <button class="fin-tab" data-col="dur"      onclick="toggleFinCol('dur')">Duration</button>
        <button class="fin-tab" data-col="ebf"      onclick="toggleFinCol('ebf')">EBF</button>
        <button class="fin-tab" data-col="collperf" onclick="toggleFinCol('collperf')">Coll. Perf.</button>
      </div>
      <button class="fin-reset-link" type="button" onclick="resetFinColumns()">Reset columns</button>
    </div>

    <div class="fin-scroll">
      <table style="border-collapse:collapse;width:100%">
        <thead class="fin-thead">
          <tr>
            <th class="fin-freeze-1" style="text-align:left;padding-left:14px;min-width:90px">CO.<br>NO.</th>
            <th class="fin-freeze-2" style="text-align:left;min-width:200px">PROJECT NAME</th>
            <th data-parentcol id="th-fin-parent" style="display:none;text-align:left;width:230px;max-width:230px;white-space:normal;line-height:1.4">PARENT PROJECT</th>
            <th style="text-align:left">PM<br>NAME</th>
            <th data-col="ov" style="text-align:right">ORDER VALUE<br>(₹ LAKHS)</th>
            <th data-col="exp" style="text-align:right">TOTAL EXPENDITURE<br>(₹ LAKHS)</th>
            <th data-col="billed" style="text-align:right">BILLED AMT<br>(₹ LAKHS)</th>
            <th data-col="col" style="text-align:right">COLLECTION<br>(₹ LAKHS)</th>
            <th data-col="bnc" style="text-align:right">BILLED, NOT<br>COLLECTED (₹ L)</th>
            <th data-col="ub" style="text-align:right">UNBILLED<br>(₹ LAKHS)</th>
            <th data-col="finpct" style="text-align:center">% EXPENDITURE OVER<br>ORDER VALUE</th>
            <th data-col="phypct" style="text-align:center">PHYSICAL<br>%</th>
            <th data-col="dur" style="text-align:center">DURATION<br>(DAYS)</th>
            <th data-col="ebf" style="text-align:center">
              <div class="ebf-mult-wrap">
                <span id="ebf-mult-label" class="ebf-mult-label">EBF = {active_ebf_mult:g}×Total Expenditure</span>
                <input type="range" id="ebf-mult-slider"
                       min="1.00" max="{GAUGE_MAX}" step="0.05" value="{active_ebf_mult}"
                       oninput="updateEbfMultiplier(this.value)">
                <span style="font-size:9px;color:{C['on_var']};font-weight:600">(₹ LAKHS)</span>
              </div>
            </th>
            <th data-col="collperf" style="text-align:center">COLL. PERF.<br>(COL−EBF)/EBF</th>
          </tr>
        </thead>
        <tbody>
          {proj_rows_html}
        </tbody>
        <tfoot class="fin-tfoot">
          <tr>
            <td class="fin-freeze-1" colspan="1" style="font-size:11px;font-weight:700;text-transform:uppercase;
                                   color:{C['dark']};letter-spacing:.05em;font-family:inherit;min-width:90px">
              Totals
            </td>
            <td class="fin-freeze-2" style="font-size:10px;font-weight:600;color:{C['on_var']};min-width:200px">
              <span id="fin-totals-scope">({sel_lob if sel_lob != "All Lines of Business" else "All LOB"})</span>
            </td>
            <td data-parentcol style="display:none"></td>
            <td></td>
            <td data-col="ov" id="total-ov" style="text-align:right;color:{C['dark']}">{total_ov_l:,.2f}</td>
            <td data-col="exp" id="total-exp" style="text-align:right;color:{C['dark']}">{total_itd_l:,.2f}</td>
            <td data-col="billed" id="total-billed" style="text-align:right;color:{C['dark']}">{total_billed_l:,.2f}</td>
            <td data-col="col" id="total-col" style="text-align:right;color:{C['dark']}">{total_col_l:,.2f}</td>
            <td data-col="bnc" id="total-bnc" style="text-align:right;color:{C['dark']}">{total_bnc_l:,.2f}</td>
            <td data-col="ub" id="total-ub" style="text-align:right;color:{C['dark']}">{total_unbilled_l:,.2f}</td>
            <td data-col="finpct" id="total-finpct" style="text-align:center;color:{C['dark']}">{avg_fin_pct:.1f}%</td>
            <td data-col="phypct" id="total-phypct" style="text-align:center;color:{C['dark']}">{avg_phy_pct:.1f}%</td>
            <td data-col="dur" id="total-dur" style="text-align:center;color:{C['dark']}">{avg_dur_str}</td>
            <td data-col="ebf" id="total-ebf" style="text-align:center;color:{C['dark']}">{total_ebf_l:,.2f}</td>
            <td data-col="collperf" style="text-align:center">
              <span id="total-perf" style="font-weight:700;color:{total_perf_color}">{total_perf_str}</span>
              <br><span style="font-size:9px;font-weight:600;color:{C['on_var']}">PORTFOLIO</span>
            </td>
          </tr>
        </tfoot>
      </table>
    </div>
  </div>

  <!-- Show More / Show Less control -->
  {f'''<div style="text-align:center;margin-bottom:20px" id="fin-showmore-wrap">
    <button type="button" id="fin-showmore-btn" onclick="finToggleShowMore()"
            style="background:#fff;border:1px solid {C['outline']};color:{C['dark']};
                   padding:9px 24px;border-radius:8px;font-size:12px;font-weight:700;
                   cursor:pointer;letter-spacing:.03em">
      ↓ Show {n_rows - INITIAL_ROWS_SHOWN} More Project{'s' if n_rows - INITIAL_ROWS_SHOWN != 1 else ''}
    </button>
  </div>''' if n_rows > INITIAL_ROWS_SHOWN else ''}

  <!-- Legend card -->
  <div style="background:{C['surface_low']};border:1px solid {C['outline']};border-radius:8px;
              padding:12px 20px;margin-bottom:20px;display:flex;gap:20px;flex-wrap:wrap;align-items:center">
    <span style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:{C['on_var']}">
      EBF Badge Key:
    </span>
    <span style="font-size:11px;font-weight:700;color:#166534;background:#f0fdf4;
                 border:1px solid #bbf7d0;padding:3px 10px;border-radius:4px">
      ✓ Collection ≥ EBF — Profit
    </span>
    <span style="font-size:11px;font-weight:700;color:#92400e;background:#fffbeb;
                 border:1px solid #fde68a;padding:3px 10px;border-radius:4px">
      ⚠ Collection 80–99% of EBF — Warning
    </span>
    <span style="font-size:11px;font-weight:700;color:#991b1b;background:#fef2f2;
                 border:1px solid #fecaca;padding:3px 10px;border-radius:4px">
      ✗ Collection &lt; 80% of EBF — Loss
    </span>
    <span style="font-size:10px;color:#9ca3af;margin-left:auto">
      Coll. Perf. = (Collection − EBF) ÷ EBF × 100
    </span>
  </div>

</div></div>

<script>
/* ── Column toggle chips (add/remove metric columns, like Consolidated Reports) ── */
var FIN_COL_ORDER = ["billed","col","ov","exp","bnc","ub","finpct","phypct","dur","ebf","collperf"];
var finActiveCols = new Set(FIN_COL_ORDER); // all shown by default
var currentEbfMult = {active_ebf_mult};
var FIN_BASE_SCOPE_LABEL = "{_fin_scope_label_js}";

function applyFinColVisibility() {{
  document.querySelectorAll('th[data-col], td[data-col]').forEach(function(el) {{
    el.style.display = finActiveCols.has(el.dataset.col) ? '' : 'none';
  }});
}}

function syncFinColButtons() {{
  document.querySelectorAll('.fin-tab').forEach(function(b) {{
    b.classList.toggle('active', finActiveCols.has(b.dataset.col));
  }});
}}

function toggleFinCol(col) {{
  if (finActiveCols.has(col)) {{
    if (finActiveCols.size > 1) finActiveCols.delete(col);
  }} else {{
    finActiveCols.add(col);
  }}
  syncFinColButtons();
  applyFinColVisibility();
  MonarchFilters.save('financials_cols', {{ cols: Array.from(finActiveCols) }});
}}

function resetFinColumns() {{
  finActiveCols = new Set(FIN_COL_ORDER);
  syncFinColButtons();
  applyFinColVisibility();
  MonarchFilters.save('financials_cols', {{ cols: Array.from(finActiveCols) }});
}}

(function() {{
  var saved = MonarchFilters.init('financials_cols');
  if (saved && Array.isArray(saved.cols) && saved.cols.length) {{
    finActiveCols = new Set(saved.cols.filter(function(c) {{ return FIN_COL_ORDER.indexOf(c) !== -1; }}));
    if (!finActiveCols.size) finActiveCols = new Set(FIN_COL_ORDER);
  }}
  syncFinColButtons();
  applyFinColVisibility();
}})();

function openProjectFolder(path) {{
  if (!path) return;
  navigator.clipboard.writeText(path).then(function() {{
    var msg = document.createElement('div');
    msg.textContent = '📋 Path copied: ' + path;
    msg.style.cssText = 'position:fixed;bottom:90px;right:28px;background:#1a2540;color:#fff;'
      + 'padding:10px 18px;border-radius:8px;font-size:13px;font-weight:600;'
      + 'z-index:9999;box-shadow:0 4px 16px rgba(0,0,0,.25);max-width:420px;word-break:break-all';
    document.body.appendChild(msg);
    setTimeout(function(){{ document.body.removeChild(msg); }}, 3000);
  }}).catch(function(){{
    alert('Network path:\\n' + path + '\\n\\nPaste this into File Explorer address bar.');
  }});
}}

/* ── Financials: Show More / Show Less ── */
var finShowingAll = false;
function finToggleShowMore() {{
  var rows = document.querySelectorAll('.fin-proj-row');
  var btn  = document.getElementById('fin-showmore-btn');
  finShowingAll = !finShowingAll;
  rows.forEach(function(tr) {{
    var idx = parseInt(tr.getAttribute('data-row-idx'), 10);
    if (idx >= {INITIAL_ROWS_SHOWN}) {{
      tr.style.display = finShowingAll ? '' : 'none';
    }}
  }});
  if (btn) {{
    btn.innerHTML = finShowingAll
      ? '↑ Show Less'
      : '↓ Show {n_rows - INITIAL_ROWS_SHOWN} More Project{"s" if n_rows - INITIAL_ROWS_SHOWN != 1 else ""}';
  }}
}}

/* ── Sub Projects toggle: flips the table from parent projects into a flat,
   read-only list of every sub-project alongside its parent's name (same
   pattern as Consolidated Reports' "Sub Projects" view). ── */
var finShowSub = false;
function toggleFinSubView() {{
  finShowSub = !finShowSub;
  document.getElementById('fin-subproj-toggle').classList.toggle('active', finShowSub);
  document.querySelectorAll('[data-parentcol]').forEach(function(el) {{
    el.style.display = finShowSub ? '' : 'none';
  }});
  var showMoreWrap = document.getElementById('fin-showmore-wrap');
  if (showMoreWrap) showMoreWrap.style.display = finShowSub ? 'none' : '';
  filterFinTable();
}}

/* ── Search box: filters fin-proj-row (or fin-child-row, in Sub Projects view)
   by project name / CO.No ── */
function filterFinTable() {{
  var q = document.getElementById('fin-search').value.toLowerCase().trim();
  var parentRows = document.querySelectorAll('.fin-proj-row');
  var childRows  = document.querySelectorAll('.fin-child-row');
  var showMoreWrap = document.getElementById('fin-showmore-wrap');
  var visible = 0;

  if (finShowSub) {{
    parentRows.forEach(function(tr) {{ tr.style.display = 'none'; }});
    childRows.forEach(function(tr) {{
      var hay = tr.getAttribute('data-search') || '';
      var show = !q || hay.indexOf(q) !== -1;
      tr.style.display = show ? '' : 'none';
      if (show) visible++;
    }});
    if (showMoreWrap) showMoreWrap.style.display = 'none';
    var badge2 = document.getElementById('fin-count-badge');
    if (badge2) badge2.textContent = 'Showing ' + visible + ' of {n_sub_rows} Sub-Projects';
    computeAndRenderFinTotals();
    return;
  }}

  childRows.forEach(function(tr) {{ tr.style.display = 'none'; }});
  parentRows.forEach(function(tr) {{
    var idx = parseInt(tr.getAttribute('data-row-idx'), 10);
    var hay = tr.getAttribute('data-search') || '';
    var show;
    if (q) {{
      show = hay.indexOf(q) !== -1;
    }} else {{
      show = finShowingAll || idx < {INITIAL_ROWS_SHOWN};
    }}
    tr.style.display = show ? '' : 'none';
    if (show) visible++;
  }});

  if (showMoreWrap) showMoreWrap.style.display = q ? 'none' : '';

  var badge = document.getElementById('fin-count-badge');
  if (badge) badge.textContent = 'Showing ' + visible + ' of {n_rows} Projects';
  computeAndRenderFinTotals();
}}

/* ── Totals row: by default reflects the whole portfolio (all parent
   projects), matching the server-rendered totals. As soon as a search term
   is typed, the totals recompute from ONLY the projects currently matching
   that search — so filtering by a name like "Gedo" totals just those rows.
   Uses each row's data-* attributes rather than the raw currency figures
   (avoiding a reload), and stays in sync with the EBF multiplier slider. ── */
function computeAndRenderFinTotals() {{
  var searchEl = document.getElementById('fin-search');
  var q = searchEl ? searchEl.value.toLowerCase().trim() : '';

  var rows = Array.from(document.querySelectorAll(finShowSub ? '.fin-child-row' : '.fin-proj-row'));
  if (q) {{
    rows = rows.filter(function(tr) {{ return tr.style.display !== 'none'; }});
  }}
  // No search term: total the full set for the current view (ignores
  // "Show More" pagination, which only hides rows visually).

  var sums = {{ ov: 0, exp: 0, billed: 0, col: 0, bnc: 0, ub: 0 }};
  var finPctSum = 0, phyPctSum = 0, durSum = 0, durCount = 0;

  rows.forEach(function(tr) {{
    sums.ov     += parseFloat(tr.dataset.ov)     || 0;
    sums.exp    += parseFloat(tr.dataset.itd)    || 0;
    sums.billed += parseFloat(tr.dataset.billed) || 0;
    sums.col    += parseFloat(tr.dataset.col)    || 0;
    sums.bnc    += parseFloat(tr.dataset.bnc)    || 0;
    sums.ub     += parseFloat(tr.dataset.ub)     || 0;
    finPctSum += parseFloat(tr.dataset.finpct) || 0;
    phyPctSum += parseFloat(tr.dataset.phypct) || 0;
    var d = tr.dataset.dur;
    if (d !== undefined && d !== '') {{
      durSum += parseFloat(d);
      durCount++;
    }}
  }});

  var n = rows.length;
  var avgFin = n ? finPctSum / n : 0;
  var avgPhy = n ? phyPctSum / n : 0;
  var avgDur = durCount ? durSum / durCount : null;
  var avgDurStr = avgDur === null ? '—'
    : (avgDur < 0 ? ('Overdue ' + Math.abs(Math.round(avgDur)) + 'd') : (Math.round(avgDur) + 'd'));

  var ebf  = Math.round(currentEbfMult * sums.exp * 100) / 100;
  var perf = ebf ? (sums.col - ebf) / ebf * 100 : 0;

  var setText = function(id, text) {{ var el = document.getElementById(id); if (el) el.textContent = text; }};
  setText('total-ov',     fmt2(sums.ov));
  setText('total-exp',    fmt2(sums.exp));
  setText('total-billed', fmt2(sums.billed));
  setText('total-col',    fmt2(sums.col));
  setText('total-bnc',    fmt2(sums.bnc));
  setText('total-ub',     fmt2(sums.ub));
  setText('total-finpct', avgFin.toFixed(1) + '%');
  setText('total-phypct', avgPhy.toFixed(1) + '%');
  setText('total-dur',    avgDurStr);
  setText('total-ebf',    fmt2(ebf));

  var perfEl = document.getElementById('total-perf');
  if (perfEl) {{
    var perfStr = (perf >= 0 ? '+' : '') + perf.toFixed(1) + '%';
    perfEl.textContent = perfStr;
    perfEl.style.color = perf >= 0 ? '#166534' : '#991b1b';
  }}

  var scopeEl = document.getElementById('fin-totals-scope');
  if (scopeEl) {{
    scopeEl.textContent = q
      ? ('Filtered — ' + n + ' project' + (n !== 1 ? 's' : '') + ' matching "' + q + '"')
      : FIN_BASE_SCOPE_LABEL;
  }}

  // Keep the Collection/EBF gauge consistent with whatever the totals row
  // is currently showing (whole portfolio, or the filtered subset).
  var ratio = ebf ? sums.col / ebf : 0;
  var ratioStr = ratio.toFixed(2) + 'x';
  var markerColor = ratio >= 1.0 ? '#166534' : '#dc2626';
  var marker = document.getElementById('gauge-marker');
  if (marker) {{
    marker.style.bottom = gaugePos(ratio).toFixed(2) + '%';
    marker.style.borderColor = markerColor;
    marker.title = 'Current: ' + ratioStr;
  }}
  var markerDot = document.getElementById('gauge-marker-dot');
  if (markerDot) markerDot.style.background = markerColor;
  var ratioStrEl = document.getElementById('gauge-ratio-str');
  if (ratioStrEl) {{ ratioStrEl.textContent = ratioStr; ratioStrEl.style.color = markerColor; }}
  var ratioBadgeEl = document.getElementById('gauge-ratio-badge');
  if (ratioBadgeEl) {{
    ratioBadgeEl.textContent = ratio >= 1.0 ? 'Profit — Collection ≥ EBF' : 'Loss — Collection < EBF';
    ratioBadgeEl.style.cssText = ratio >= 1.0
      ? 'display:inline-block;padding:5px 18px;border-radius:6px;font-size:12px;font-weight:700;margin-top:8px;background:#f0fdf4;color:#166534'
      : 'display:inline-block;padding:5px 18px;border-radius:6px;font-size:12px;font-weight:700;margin-top:8px;background:#fef2f2;color:#991b1b';
  }}
  var avgPerfEl = document.getElementById('gauge-avg-perf');
  if (avgPerfEl) {{
    var avgCollRate = sums.ov ? (sums.col / sums.ov * 100) : 0;
    avgPerfEl.textContent = 'Avg Collection Rate: ' + avgCollRate.toFixed(1) + '% of Order Value';
  }}

  var oor = document.getElementById('gauge-oor');
  if (oor) {{
    if (ratio < GAUGE_MIN) {{
      oor.style.color = '#991b1b';
      oor.textContent = 'Below ' + GAUGE_MIN.toFixed(2) + 'x — pinned to scale floor';
    }} else if (ratio > GAUGE_MAX) {{
      oor.style.color = '#166534';
      oor.textContent = 'Above ' + GAUGE_MAX.toFixed(2) + 'x — pinned to scale ceiling';
    }} else {{
      oor.textContent = '';
    }}
  }}

  var subEl = document.getElementById('gauge-sub');
  if (subEl) {{
    if (ratio >= 1.0) {{
      var surplus = (sums.col - ebf) / 1e5;
      subEl.textContent = 'Surplus of ₹' + surplus.toFixed(1) + 'L above EBF target';
    }} else {{
      var deficit = (ebf - sums.col) / 1e5;
      subEl.textContent = 'Deficit of ₹' + deficit.toFixed(1) + 'L below EBF target';
    }}
  }}

  var covPct = Math.min(ratio * 100, 100);
  var covPctEl = document.getElementById('cov-pct');
  if (covPctEl) {{
    covPctEl.textContent = covPct.toFixed(1) + '%';
    covPctEl.style.color = ratio >= 1 ? '#166534' : '#991b1b';
  }}
  var covBarEl = document.getElementById('cov-bar');
  if (covBarEl) {{
    covBarEl.style.width = covPct.toFixed(1) + '%';
    covBarEl.style.background = ratio >= 1 ? '#166534' : '#dc2626';
  }}
  var covTargetEl = document.getElementById('cov-target');
  if (covTargetEl) covTargetEl.textContent = 'EBF target: ₹' + (ebf / 1e5).toFixed(1) + 'L';
}}

/* ── EBF multiplier vertical scroller logic ──
   Recomputes EBF, Coll.Perf, badges, totals row, and the gauge,
   entirely client-side, for any multiplier between 1.00x and 1.35x. */
var GAUGE_MIN = 1.00, GAUGE_MAX = 1.35;

function gaugePos(val) {{
  var v = Math.max(GAUGE_MIN, Math.min(GAUGE_MAX, val));
  return (v - GAUGE_MIN) / (GAUGE_MAX - GAUGE_MIN) * 100;
}}

function fmt2(n) {{
  return n.toLocaleString('en-IN', {{minimumFractionDigits: 2, maximumFractionDigits: 2}});
}}

function updateEbfMultiplier(multStr) {{
  var mult = parseFloat(multStr);
  currentEbfMult = mult;

  // Keep the "Set" input in sync with whatever the slider is showing, so
  // clicking Set persists exactly the value the slider picked.
  var ebfInput = document.getElementById('ebf-mult-input');
  if (ebfInput && document.activeElement !== ebfInput) {{
    ebfInput.value = mult.toFixed(2);
  }}

  var exportLink = document.getElementById('export-excel-link');
  if (exportLink) {{
    var url = new URL(exportLink.href, window.location.origin);
    url.searchParams.set('ebf_mult', mult.toFixed(2));
    exportLink.href = url.toString();
  }}

  document.getElementById('ebf-mult-label').textContent = 'EBF = ' + mult.toFixed(2) + '×Total Expenditure';

  // Per-row EBF / Coll.Perf badges always reflect every project in the
  // table (not just the currently filtered/visible ones) so the figures
  // are correct as soon as a search is cleared or a filter changes.
  var rows = document.querySelectorAll('.fin-wrap tbody tr[data-itd]');

  rows.forEach(function(tr) {{
    var itd = parseFloat(tr.dataset.itd);
    var col = parseFloat(tr.dataset.col);

    var ebf = Math.round(mult * itd * 100) / 100;
    var perf = ebf ? (col - ebf) / ebf * 100 : 0;

    var ebfBadge = tr.querySelector('.ebf-badge');
    var perfBadge = tr.querySelector('.perf-badge');

    ebfBadge.textContent = fmt2(ebf);
    if (col >= ebf) {{
      ebfBadge.style.cssText = "display:inline-block;padding:3px 10px;border-radius:4px;font-size:11px;font-weight:700;color:#166534;background:#f0fdf4;border:1px solid #bbf7d0";
    }} else if (col >= ebf * 0.80) {{
      ebfBadge.style.cssText = "display:inline-block;padding:3px 10px;border-radius:4px;font-size:11px;font-weight:700;color:#92400e;background:#fffbeb;border:1px solid #fde68a";
    }} else {{
      ebfBadge.style.cssText = "display:inline-block;padding:3px 10px;border-radius:4px;font-size:11px;font-weight:700;color:#991b1b;background:#fef2f2;border:1px solid #fecaca";
    }}

    var perfStr = (perf >= 0 ? '+' : '') + perf.toFixed(1) + '%';
    perfBadge.textContent = perfStr;
    if (perf >= 0) {{
      perfBadge.style.cssText = "font-weight:700;color:#166534;background:#f0fdf4;padding:3px 8px;border-radius:4px;font-size:11px";
    }} else {{
      perfBadge.style.cssText = "font-weight:700;color:#991b1b;background:#fef2f2;padding:3px 8px;border-radius:4px;font-size:11px";
    }}
  }});

  // Totals row + gauge: recomputed from whatever the search filter currently
  // shows (or the whole portfolio, if there's no active search).
  computeAndRenderFinTotals();
}}

// ── Manual "Set" input <-> slider two-way sync ──
// Typing a value into the Set box also moves the slider (if in range) and
// live-previews the table/gauge, same as dragging the slider does.
(function() {{
  var ebfInput = document.getElementById('ebf-mult-input');
  var slider   = document.getElementById('ebf-mult-slider');
  if (!ebfInput) return;
  ebfInput.addEventListener('input', function() {{
    var v = parseFloat(this.value);
    if (isNaN(v)) return;
    if (slider) {{
      var min = parseFloat(slider.min), max = parseFloat(slider.max);
      slider.value = Math.min(Math.max(v, min), max);
    }}
    updateEbfMultiplier(v);
  }});
}})();
</script>

<!-- FAB -->
<button onclick="window.location='/data_entry'" title="Create Project"
        style="position:fixed;bottom:28px;right:28px;width:52px;height:52px;
               background:{C['dark']};color:#fff;border:none;border-radius:50%;
               font-size:26px;cursor:pointer;box-shadow:0 4px 16px rgba(0,0,0,.24)">＋</button>

</body></html>"""
    return html

# ═══════════════════════════════════════════════════════════════
#  PAGE 5 — VARIATIONS
# ═══════════════════════════════════════════════════════════════
def export_variations_pdf_bytes() -> bytes:
    db_vars = fetch_all_variations()
    rows, total_var = [], 0.0
    for v in db_vars:
        amt = _n(v.get("variation_amt")); total_var += amt
        rows.append([v.get("child_co_no",""), v.get("child_name",""), v.get("parent_name",""),
                     f"{_n(v.get('original_value')):,.0f}", f"{amt:,.0f}",
                     f"{_n(v.get('revised_value')):,.0f}", v.get("variation_date","")])
    return _build_table_pdf(
        "Parent Project Details",
        f"{len(db_vars)} variation records",
        ["Child CO.No","Child Name","Parent","Original (₹)","Variation (₹)","Revised (₹)","Date"],
        rows, footer_label="Net Variation", footer_value=total_var)


def variations(alert_html="", prefill=None, username="Admin") -> str:
    stats       = fetch_summary_stats()
    db_projects = fetch_all_projects()
    db_vars     = fetch_all_variations()
    proj_rows   = fetch_projects_for_select()

    net_v   = stats["net_variation"]
    net_fmt = f"₹{net_v/1e6:.1f}M" if net_v >= 1e6 else f"₹{net_v:,.0f}"

    cutoff_iso     = (datetime.now() - timedelta(days=90)).isoformat()
    parents_last_q = sum(
        1 for p in db_projects
        if p.get("project_type") != "child"
        and (p.get("created_at") or "") >= cutoff_iso
    )
    total_ov_db = sum(_n(p.get("order_value")) for p in db_projects)
    avg_children = f"{stats['total_children'] / max(stats['total_parents'], 1):.1f} avg/parent"
    net_var_pct  = (f"{abs(net_v) / max(total_ov_db, 1) * 100:.1f}% of Order Value"
                    if total_ov_db else "—")

    # ── SWAPPED: now uses rich financials-style expand with KPI cards ──
    parent_list_v = [p for p in db_projects if p.get("project_type", "parent") != "child"]
    child_list_v  = [p for p in db_projects if p.get("project_type", "") == "child"]

    # ── Only show parents that actually have at least one sub-project.
    # A "parent" record with zero linked children isn't part of any real
    # hierarchy, so it's excluded from this page entirely (both the table
    # below and the click-through popup). ──
    def _parent_has_kids(p):
        p_co   = (p.get("co_no") or "").strip()
        p_name = (p.get("project_name") or "").strip().lower()
        return any(
            (p_co and (c.get("parent_id") or "").strip() == p_co)
            or (c.get("parent_project") or "").strip().lower() == p_name
            for c in child_list_v
        )
    parent_list_v = [p for p in parent_list_v if _parent_has_kids(p)]

    def _view_billed_total(project):
        billing_rows = fetch_billing_entries(project.get("id"))
        billing_total = sum(_n(row.get("billed_amt")) for row in billing_rows)
        return billing_total or _n(project.get("billed_amount"))

    # ── Total Expense (ITD Cost) for a project: use Data Entry's own
    # "Total ITD Cost (₹)" field (projects.itd_cost) if it's been recorded;
    # otherwise fetch the live sum from the Expenditures tab (erp_updates).
    # If NEITHER has any data, show 0 — this must never fall back to
    # billed_amount, since that's the Billed figure (amount invoiced), a
    # completely different thing from Expense (cost incurred). ──
    def _view_expense_total(project):
        de_itd = _n(project.get("itd_cost"))
        if de_itd:
            return de_itd
        erp_sum = sum(_n(v) for v in fetch_erp_totals(project.get("co_no")).values())
        return erp_sum  # 0 if Expenditures has nothing either

    hier_rows = ""
    for par in parent_list_v:
        co   = par.get("co_no")           or "—"
        name = par.get("project_name")    or "—"
        pm   = par.get("project_manager") or "—"
        # ── Find this parent's children, needed for the "N sub-projects"
        # badge, the expandable child rows below, and the popup. ──
        par_co   = (par.get("co_no") or "").strip()
        kids     = [c for c in child_list_v
                    if (par_co and (c.get("parent_id") or "").strip() == par_co)
                    or (c.get("parent_project") or "").strip().lower() == name.strip().lower()]
        has_kids = len(kids) > 0

        # ── This row shows the PARENT PROJECT's own details — Order Value,
        # ITD Billed, and Collection exactly as maintained in Data Entry for
        # this parent record, NOT rolled up from its sub-projects. Rolled-up
        # sub-project totals are shown separately: in the expandable child
        # rows below, and in the popup's summary cards when you click in. ──
        ov     = _rev_ov(par)
        rev_ov = ov
        itd    = _view_billed_total(par)
        col    = _n(par.get("collection_till_date"))

        # compute duration in days from today to end date (prefer extended_end_date if present)
        duration_days = ""
        end_str = (par.get("extended_end_date") or par.get("end_date") or "")
        if end_str:
          try:
            end_dt = datetime.fromisoformat(str(end_str)).date()
          except Exception:
            try:
              end_dt = datetime.strptime(str(end_str)[:10], "%Y-%m-%d").date()
            except Exception:
              end_dt = None
          if end_dt:
            duration_days = (end_dt - datetime.now().date()).days
        # prepare display: ONLY show two states — overdue (red) or days remaining (green).
        if isinstance(duration_days, int):
            if duration_days < 0:
                dur_display = f"Overdue {abs(duration_days)} days"
                dur_badge_style = (
                    f"display:inline-block;padding:6px 10px;border-radius:999px;"
                    f"background:#fff1f0;color:{C['error']};border:1px solid #fecaca;"
                    f"font-size:12px;font-weight:700;"
                )
            else:
                dur_display = f"{duration_days} days"
                dur_badge_style = (
                    f"display:inline-block;padding:6px 10px;border-radius:999px;"
                    f"background:#f0fdf4;color:#166534;border:1px solid #bbf7d0;"
                    f"font-size:12px;font-weight:700;"
                )
        else:
            # no date available — show nothing
            dur_display = ""
            dur_badge_style = ""
        ebf  = round(1.35 * itd, 2)                  # Col AA
        perf = (col / rev_ov * 100) if rev_ov else 0 # Col AB
        pid  = par.get("id", "")
        st   = par.get("project_status") or "Active"

        # ── New main-table columns: Order Date, Order Value, Variation, Time Extension ──
        order_date_raw = (par.get("order_issue_date") or "").strip()
        if order_date_raw:
            try:
                order_date_disp = datetime.fromisoformat(str(order_date_raw)[:10]).strftime("%d-%b-%Y")
            except Exception:
                order_date_disp = str(order_date_raw)[:10]
        else:
            order_date_disp = "—"
        order_value_disp = _n(par.get("order_value"))
        variation_disp   = _n(par.get("variation_amt"))
        ext_months_val   = par.get("extension_months")
        if ext_months_val:
            try:
                _emf = float(ext_months_val)
                ext_months_disp = str(int(_emf)) if _emf == int(_emf) else str(_emf)
            except Exception:
                ext_months_disp = str(ext_months_val)
        else:
            ext_months_disp = "—"

        ratio_ebf = ebf / itd if itd else 0
        if ratio_ebf >= 1.30:   ebf_cls = "background:#f0fdf4;color:#166534;border:1px solid #bbf7d0"
        elif ratio_ebf >= 1.10: ebf_cls = "background:#fffbeb;color:#92400e;border:1px solid #fde68a"
        else:                   ebf_cls = "background:#fef2f2;color:#991b1b;border:1px solid #fecaca"

        if st in ("Cancelled","Hold"):            st_cls = "background:#ffdad6;color:#ba1a1a"
        elif st in ("Closed","Execution Closed"): st_cls = f"background:{C['surface_high']};color:{C['on_var']}"
        else:                                     st_cls = "background:#fde68a;color:#92400e"

        pp_search_val = (f"{pid} {co} {name} {st} {pm}".lower()
                         .replace('&', '&amp;').replace('"', '&quot;').replace("'", "&#39;"))
        pp_pm_val = (pm or "").strip().lower().replace('&', '&amp;').replace('"', '&quot;').replace("'", "&#39;")
        hier_rows += f"""
        <tr class="parent-fin-row" data-pid="{pid}" data-open="0" data-search="{pp_search_val}" data-pm="{pp_pm_val}"
            style="cursor:{'pointer' if has_kids else 'default'};"
            {'onclick="openChildModal(\''+str(pid)+'\')"' if has_kids else ''}>
          <td style="padding:14px 16px">
            <div style="display:flex;align-items:center;gap:10px">
              <span class="expand-icon" style="font-size:18px;color:{C['on_var']};
                    transition:transform .2s;display:inline-block;min-width:20px;text-align:center">
                {'›' if has_kids else '—'}
              </span>
              <div>
                <p style="font-weight:700;color:{C['dark']};font-size:14px">{co}: {name}</p>
                <div style="display:flex;align-items:center;gap:8px;margin-top:4px">
                  <span style="padding:2px 8px;border-radius:3px;font-size:10px;
                               font-weight:700;text-transform:uppercase;{st_cls}">{st}</span>
                  <span style="font-size:11px;color:{C['on_var']}">{pm}</span>
                  {f'<span style="font-size:10px;background:{C["dark"]};color:#fff;padding:1px 7px;border-radius:10px;font-weight:700">{len(kids)} sub-project{"s" if len(kids)>1 else ""}</span>' if has_kids else ''}
                </div>
              </div>
            </div>
          </td>
          <td style="text-align:right;padding:14px 16px">
            <span style="{dur_badge_style}">{dur_display}</span>
          </td>
          <td style="text-align:right;font-family:'Courier New',monospace;font-size:13px;padding:14px 16px;font-weight:600">{order_date_disp}</td>
          <td style="text-align:right;font-family:'Courier New',monospace;font-size:13px;padding:14px 16px;font-weight:600">{order_value_disp:,.0f}</td>
          <td style="text-align:right;font-family:'Courier New',monospace;font-size:13px;padding:14px 16px;font-weight:600">{variation_disp:,.0f}</td>
          <td style="text-align:right;font-family:'Courier New',monospace;font-size:13px;padding:14px 16px;font-weight:600">{ext_months_disp}</td>
        </tr>"""

        if has_kids:
            for ch in kids:
                ch_co   = ch.get("co_no",          "—")
                ch_name = ch.get("project_name",   "—")
                ch_pm   = ch.get("project_manager","—")
                ch_st   = ch.get("project_status", "Active")
                ch_ov   = _rev_ov(ch)
                ch_itd  = _view_billed_total(ch)
                ch_col  = _n(ch.get("collection_till_date"))
                ch_ebf  = round(1.35 * ch_itd, 2)
                ch_perf = (ch_col / ch_ov * 100) if ch_ov else 0

                ch_budget    = ch_ov
                ch_expense   = _view_expense_total(ch)
                ch_remaining = ch_budget - ch_expense
                ch_exp_pct   = (ch_expense / ch_budget * 100) if ch_budget else 0

                if ch_exp_pct >= 100:
                    warn_badge = (
                        f'<span style="display:inline-flex;align-items:center;gap:4px;'
                        f'background:#fef2f2;color:#991b1b;border:1px solid #fca5a5;'
                        f'padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700">🚨 CRITICAL — Over Budget</span>'
                    )
                    exp_pct_style   = "color:#991b1b;font-weight:800"
                    bar_color       = "#dc2626"
                    remaining_style = "color:#991b1b;font-weight:700"
                elif ch_exp_pct >= 80:
                    warn_badge = (
                        f'<span style="display:inline-flex;align-items:center;gap:4px;'
                        f'background:#fff7ed;color:#c2410c;border:1px solid #fdba74;'
                        f'padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700">⚠ WARNING — High Spend</span>'
                    )
                    exp_pct_style   = "color:#c2410c;font-weight:700"
                    bar_color       = "#f97316"
                    remaining_style = "color:#c2410c;font-weight:600"
                else:
                    warn_badge      = ""
                    exp_pct_style   = f"color:{C['dark']};font-weight:600"
                    bar_color       = "#166534"
                    remaining_style = "color:#166534;font-weight:600"

                bar_width = min(ch_exp_pct, 100)

                if ch_st in ("Cancelled","Hold"):            ch_st_cls = "background:#ffdad6;color:#ba1a1a"
                elif ch_st in ("Closed","Execution Closed"): ch_st_cls = f"background:{C['surface_high']};color:{C['on_var']}"
                else:                                         ch_st_cls = "background:#e0f2fe;color:#0369a1"

                hier_rows += f"""
                <tr class="child-fin-row child-of-{pid}" style="display:none;background:#f8faff;">
                  <td colspan="6" style="padding:0 0 0 56px;">
                    <div style="border-left:3px solid {C['mid']};margin:6px 16px 6px 0;padding:14px 16px;
                                background:#fff;border-radius:0 8px 8px 0;
                                border:1px solid {C['outline']};border-left:3px solid {C['mid']};">
                      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;">
                        <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
                          <span style="font-size:13px;font-weight:700;color:{C['dark']}">{ch_co}: {ch_name}</span>
                          <span style="padding:2px 8px;border-radius:3px;font-size:10px;font-weight:700;text-transform:uppercase;{ch_st_cls}">{ch_st}</span>
                          <span style="font-size:11px;color:{C['on_var']}">PM: {ch_pm}</span>
                          {warn_badge}
                        </div>
                      </div>
                      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:12px;">
                        <div style="background:{C['surface_low']};border-radius:6px;padding:10px 12px;border:1px solid {C['outline']};">
                          <div style="font-size:9px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['on_var']};margin-bottom:4px;">Budget (Order Val)</div>
                          <div style="font-size:15px;font-weight:700;color:{C['dark']};">₹{ch_budget:,.0f}</div>
                        </div>
                        <div style="background:{C['surface_low']};border-radius:6px;padding:10px 12px;border:1px solid {C['outline']};">
                          <div style="font-size:9px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['on_var']};margin-bottom:4px;">Expense (ITD Cost)</div>
                          <div style="font-size:15px;{exp_pct_style};">₹{ch_expense:,.0f}</div>
                        </div>
                        <div style="background:{C['surface_low']};border-radius:6px;padding:10px 12px;border:1px solid {C['outline']};">
                          <div style="font-size:9px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['on_var']};margin-bottom:4px;">Remaining Budget</div>
                          <div style="font-size:15px;{remaining_style};">₹{ch_remaining:,.0f}</div>
                        </div>
                        <div style="background:{C['surface_low']};border-radius:6px;padding:10px 12px;border:1px solid {C['outline']};">
                          <div style="font-size:9px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['on_var']};margin-bottom:4px;">Expense %</div>
                          <div style="font-size:15px;{exp_pct_style};">{ch_exp_pct:.1f}%</div>
                        </div>
                      </div>
                      <div>
                        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">
                          <span style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:{C['on_var']};">Budget Utilization</span>
                          <span style="font-size:11px;font-weight:700;{exp_pct_style};">{ch_exp_pct:.1f}% used</span>
                        </div>
                        <div style="background:#e5e7eb;border-radius:99px;height:8px;overflow:hidden;">
                          <div style="height:100%;width:{bar_width:.1f}%;background:{bar_color};border-radius:99px;transition:width .4s ease;"></div>
                        </div>
                      </div>
                      <div style="display:flex;gap:24px;margin-top:12px;padding-top:10px;border-top:1px solid {C['outline']};flex-wrap:wrap;">
                         <div><span style="font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']};">Total Expenditure Billed</span>
                          <span style="font-size:12px;font-weight:600;color:{C['dark']};margin-left:6px;font-family:'Courier New',monospace;">₹{ch_itd:,.0f}</span></div>
                        <div><span style="font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']};">Collection</span>
                          <span style="font-size:12px;font-weight:600;color:{C['dark']};margin-left:6px;font-family:'Courier New',monospace;">₹{ch_col:,.0f}</span></div>
                        <div><span style="font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']};">EBF (1.35×)</span>
                          <span style="font-size:12px;font-weight:600;color:{C['dark']};margin-left:6px;font-family:'Courier New',monospace;">₹{ch_ebf:,.0f}</span></div>
                        <div><span style="font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']};">Coll. Perf.</span>
                          <span style="font-size:12px;font-weight:600;color:{C['dark']};margin-left:6px;">{ch_perf:.1f}%</span></div>
                      </div>
                    </div>
                  </td>
                </tr>"""

    if not hier_rows:
      hier_rows = f'<tr><td colspan="6" style="padding:20px;text-align:center;color:#aaa">No projects. <a href="/data_entry">Add via Data Entry →</a></td></tr>'

    var_rows = ""
    for v in db_vars:
        amt     = _n(v.get("variation_amt"))
        orig    = _n(v.get("original_value"))
        revised = _n(v.get("revised_value"))
        if amt > 0:   amt_style = f"color:{C['error']};font-weight:700"; amt_str = f"+{_fmt_inr(amt)}"
        elif amt < 0: amt_style = "color:#166534;font-weight:700";        amt_str = _fmt_inr(amt)
        else:         amt_style = f"color:{C['on_var']}";                 amt_str = "0"
        var_rows += f"""
        <tr>
          <td>
            <p style="font-weight:700;color:{C['dark']};font-size:13px">{v.get('child_co_no','')}: {v.get('child_name','')}</p>
            <p style="font-size:11px;color:{C['on_var']}">Parent: {v.get('parent_name','')}</p>
          </td>
          <td style="text-align:right;font-family:'Courier New',monospace">{_fmt_inr(orig)}</td>
          <td style="text-align:right;font-family:'Courier New',monospace;{amt_style}">{amt_str}</td>
          <td style="text-align:right;font-family:'Courier New',monospace;font-weight:700">{_fmt_inr(revised)}</td>
          <td style="font-family:'Courier New',monospace;font-size:13px;color:{C['on_var']}">{v.get('variation_date','')}</td>
          <td style="text-align:right">
            <form method="post" action="/delete_variation" style="display:inline">
              <input type="hidden" name="var_id" value="{v['id']}">
              <button class="del-btn">🗑</button>
            </form>
          </td>
        </tr>"""
    if not var_rows:
        var_rows = '<tr><td colspan="6" style="padding:20px;text-align:center;color:#aaa">No variations. Use "+ New Variation".</td></tr>'

    proj_sel_opts = "".join(
        f'<option value="{r[0]}">{r[1] or ""} | {r[2] or ""}</option>'
        for r in proj_rows
    ) or '<option value="">— No projects —</option>'
    # Build JSON for modal
    import json
    modal_data = {}
    for par in parent_list_v:
        pid_key = str(par.get("id", ""))
        par_co  = (par.get("co_no") or "").strip()
        kids    = [c for c in child_list_v
                   if (par_co and (c.get("parent_id") or "").strip() == par_co)
                   or (c.get("parent_project") or "").strip().lower()
                   == (par.get("project_name") or "").strip().lower()]
        # ── Total Expense per sub-project: sourced from the Data Entry tab's
        # "Total ITD Cost (₹)" field (projects.itd_cost) first, since that is
        # the value the user maintains directly on that project. Falls back
        # to the live Expenditures/erp_updates sum only if Data Entry has no
        # itd_cost recorded yet for that sub-project; if Expenditures has
        # nothing either, this shows 0 rather than borrowing billed_amount
        # (a different field — Expense vs. Billed are kept independent). ──
        child_expense = {}
        for c in kids:
            cid = c.get("id")
            child_expense[cid] = _view_expense_total(c)

        # ── Sub-project's OWN ITD Billed — sourced from the Data Entry tab's
        # "Up to date Billed Amount (₹)" field (projects.billed_amount) first,
        # since that is the value the user maintains directly on that project.
        # Falls back to the live billing_collection sum only if Data Entry has
        # no billed_amount recorded yet for that sub-project. This keeps the
        # popup's per-project figure consistent with what's shown/edited in
        # Data Entry, and with the header total below (which also reads
        # billed_amount / itd_cost from Data Entry). ──
        child_billed = {}
        for c in kids:
            cid = c.get("id")
            child_billed[cid] = _view_billed_total(c)

        # ── ITD Billed / Order Value / Collection shown in the popup's TOP
        # summary line are the PARENT PROJECT's own values, exactly as
        # maintained directly in Data Entry for this parent record — NOT
        # rolled up from its sub-projects. The cards just below (Total Rev.
        # OV, ITD Billed, Collection) are the separate rolled-up totals from
        # the sub-projects. Keeping these independent lets the person compare
        # "what the parent record itself says" against "what its
        # sub-projects add up to". ──
        parent_own_order_value = _n(par.get("order_value"))
        parent_own_itd_billed  = _view_billed_total(par)
        parent_own_collection  = _n(par.get("collection_till_date"))

        table_itd_billed = (
            sum(child_billed.values())
            if kids else parent_own_itd_billed
        )
        table_order_value = (
            sum(_n(c.get("revised_order_value")) or
                (_n(c.get("order_value")) + _n(c.get("variation_amt")))
                for c in kids)
            if kids else parent_own_order_value
        )
        table_collection = (
            sum(_n(c.get("collection_till_date")) for c in kids)
            if kids else parent_own_collection
        )

        rolled_itd = (sum(child_expense.values())
                      if kids else _n(par.get("itd_cost")))
        modal_data[pid_key] = {
            "parent_name": par.get("project_name", "—"),
            "parent_co":   par.get("co_no", "—"),
            "order_value": parent_own_order_value,
            "itd_billed":  parent_own_itd_billed,
            "collection":  parent_own_collection,
            "total_rev_ov":     table_order_value,
            "total_itd_billed": table_itd_billed,
            "total_collection": table_collection,
            "children": [
                {
                    "co_no":        c.get("co_no", "—"),
                    "name":         c.get("project_name", "—"),
                    "pm":           c.get("project_manager", "—"),
                    "status":       c.get("project_status", "—"),
                    "order_value":  _n(c.get("order_value")),
                    "revised_ov":   (_n(c.get("revised_order_value")) or (_n(c.get("order_value")) + _n(c.get("variation_amt")))),
                    "variation":    _n(c.get("variation_amt")),
                    "itd_cost":     child_expense.get(c.get("id"), 0),
                    "billed":       child_billed.get(c.get("id"), 0),
                    "collection":   _n(c.get("collection_till_date")),
                    "phys_pct":     _n(c.get("physical_completion")),
                    "ebf":          round(1.35 * child_billed.get(c.get("id"), 0), 2),
                    "baseline_eac": _n(c.get("baseline_eac")),
                    "rev_eac":      _n(c.get("rev_latest_eac")),
                    # Margin % — Parent Project Details popup ONLY: computed live as
                    # (1 - Rev. EAC / Revised OV) x 100, using the same Revised OV and
                    # Rev. EAC already shown on this card, instead of the cached
                    # rev_latest_margin column (which can go stale relative to this
                    # view). Other tabs/pages are unaffected and keep using the
                    # stored rev_latest_margin as before.
                    "margin": (
                        round((1 - _n(c.get("rev_latest_eac")) /
                               (_n(c.get("revised_order_value")) or (_n(c.get("order_value")) + _n(c.get("variation_amt"))))) * 100, 2)
                        if (_n(c.get("revised_order_value")) or (_n(c.get("order_value")) + _n(c.get("variation_amt"))))
                        else 0
                    ),
                }
                for c in kids
            ]
        }
    modal_data_json = json.dumps(modal_data) 
    html  = _html_head("Parent Project Details") + "<body>"
    html += _sidebar("variations") + _topbar("Parent Project Details", username)
    html += f"""
<style>
.parent-fin-row:hover td{{background:rgba(0,14,36,.03)}}
.var-panel{{position:fixed;inset-y:0;right:0;width:420px;background:#fff;border-left:1px solid {C['outline']};box-shadow:-4px 0 24px rgba(0,0,0,.12);z-index:400;transform:translateX(100%);transition:transform .3s ease;display:flex;flex-direction:column}}
.var-panel.open{{transform:translateX(0)}}
.overlay{{position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:399;display:none;backdrop-filter:blur(1px)}}
.overlay.show{{display:block}}
@keyframes fadeSlide{{from{{opacity:0;transform:translateY(-4px)}}to{{opacity:1;transform:translateY(0)}}}}
</style>
<div class="main"><div class="canvas" style="max-width:1360px">
  {alert_html}
  <div style="display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:24px">
    <div>
      <h2 style="font-size:28px;font-weight:700;color:{C['dark']}">Parent Project Details</h2>
      <p style="font-size:14px;color:#6b7280;margin-top:4px">Live hierarchy and financial deviation analysis.</p>
    </div>
    <div style="display:flex;gap:10px">
      <a href="/export_excel/variations"
         style="display:inline-flex;align-items:center;gap:6px;padding:8px 16px;
                background:{C['dark']};color:#fff;border-radius:7px;font-size:12px;
                font-weight:700;text-decoration:none;letter-spacing:.04em"
      >⬇ Export Excel</a>
    </div>
  </div>

  <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:28px">
    {" ".join(f'<div style="background:{C["surface_low"]};border:1px solid {C["outline"]};border-radius:8px;padding:14px 16px"><div style="font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C["on_var"]};margin-bottom:8px">{lbl}</div><div style="display:flex;align-items:flex-end;justify-content:space-between"><span style="font-size:26px;font-weight:800;color:{vc}">{vl}</span><span style="font-size:11px;font-weight:600;color:{sc}">{sub}</span></div></div>'
      for lbl,vl,vc,sub,sc in [
        ("Total Parent Projects",  str(stats['total_parents']),  C['dark'], f"+{parents_last_q} Last 90 Days", "#92400e"),
        ("Total Child Projects",   str(stats['total_children']), C['dark'], avg_children,                      C['on_var']),
        
      ])}
  </div>

  <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;overflow:hidden;margin-bottom:24px">
    <div style="padding:12px 18px;border-bottom:1px solid {C['outline']};background:{C['surface_low']};
                display:flex;align-items:center;justify-content:space-between;gap:16px;flex-wrap:wrap">
      <span style="font-size:11px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['dark']}">Project Hierarchy &amp; Financial Roll-up</span>
      <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
        <select id="pp-pm-filter" onchange="filterParentRows()"
                style="padding:8px 12px;border:1px solid {C['outline']};border-radius:6px;font-size:12px;
                       background:#fff;color:{C['dark']};outline:none;cursor:pointer;min-width:170px">
          <option value="">All PMs</option>
          {"".join(f'<option value="{pmn.lower()}">{pmn}</option>' for pmn in get_project_managers())}
        </select>
        <div style="position:relative;min-width:280px;flex:0 0 auto">
          <input type="text" id="pp-search" placeholder="Search Project ID, CO No., Name or Status…"
                 oninput="filterParentRows()"
                 style="width:100%;padding:8px 12px 8px 30px;border:1px solid {C['outline']};
                        border-radius:6px;font-size:12px;background:#fff;outline:none;color:{C['dark']}">
          <span style="position:absolute;left:10px;top:50%;transform:translateY(-50%);font-size:12px;color:{C['on_var']}">🔍</span>
        </div>
      </div>
    </div>
    <div style="overflow-x:auto">
      <table style="width:100%;border-collapse:collapse;font-size:13px;table-layout:fixed">
        <thead><tr style="background:{C['surface_low']};border-bottom:1px solid {C['outline']}">
          <th style="text-align:left;padding-left:16px;width:30%">CO.NO. &amp; Project Name</th>
          <th style="text-align:right;width:10%">Duration Status</th>
          <th style="text-align:right;width:14%">Order Date</th>
          <th style="text-align:right;width:16%">Order Value (₹)</th>
          <th style="text-align:right;width:16%">Variation (₹)</th>
          <th style="text-align:right;width:14%">Time Extension (Months)</th>
        </tr></thead>
        <tbody id="pp-tbody">{hier_rows}</tbody>
      </table>
      <div id="pp-no-match" style="display:none;padding:20px;text-align:center;color:#aaa">No projects match your search.</div>
    </div>
  </div>

  <!-- Legend for warning colors -->
  <div style="display:flex;gap:12px;align-items:center;margin-top:14px;margin-bottom:4px;flex-wrap:wrap;">
    <span style="font-size:11px;font-weight:700;text-transform:uppercase;color:{C['on_var']};letter-spacing:.04em;">Child Project Status:</span>
    <span style="display:inline-flex;align-items:center;gap:5px;font-size:11px;font-weight:600;color:#166534;background:#f0fdf4;border:1px solid #bbf7d0;padding:2px 10px;border-radius:4px;">✓ &lt;80% — On Track</span>
    <span style="display:inline-flex;align-items:center;gap:5px;font-size:11px;font-weight:600;color:#c2410c;background:#fff7ed;border:1px solid #fdba74;padding:2px 10px;border-radius:4px;">⚠ 80–99% — High Spend</span>
    <span style="display:inline-flex;align-items:center;gap:5px;font-size:11px;font-weight:600;color:#991b1b;background:#fef2f2;border:1px solid #fca5a5;padding:2px 10px;border-radius:4px;">🚨 ≥100% — Over Budget</span>
    <span style="display:inline-flex;align-items:center;gap:6px;font-size:11px;font-weight:600;color:{C['on_var']};background:#f8fafc;border:1px solid {C['outline']};padding:2px 10px;border-radius:4px;">Duration Status = If overdue → Overdue N days; else → N days</span>
    <span style="display:inline-flex;align-items:center;gap:6px;font-size:11px;font-weight:600;color:#166534;background:#f0fdf4;border:1px solid #bbf7d0;padding:2px 10px;border-radius:4px;">Green = Future/Remaining days</span>
    <span style="display:inline-flex;align-items:center;gap:6px;font-size:11px;font-weight:600;color:{C['error']};background:#fff1f0;border:1px solid #fecaca;padding:2px 10px;border-radius:4px;">Red = Overdue</span>
  </div>
  <div style="text-align:right;font-size:11px;font-style:italic;color:#9ca3af;margin-top:6px">
    Figures in ₹ · Click › on a parent row to expand child project details
  </div>
</div></div>

<div class="var-panel" id="var-panel">
  <div style="padding:18px 20px;border-bottom:1px solid {C['outline']};background:{C['surface_low']};display:flex;align-items:center;justify-content:space-between">
    <span style="font-size:16px;font-weight:700;color:{C['dark']}">New Variation Entry</span>
    <button onclick="closeVarPanel()" style="background:none;border:none;font-size:20px;cursor:pointer;color:{C['on_var']}">✕</button>
  </div>
  <div style="flex:1;padding:22px 20px;overflow-y:auto">
    <form method="post" action="/save_variation" id="var-form">
      {" ".join([
        _field_html("Child CO. No.",   '<input type="text" name="child_co_no" placeholder="e.g. C-24-03">'),
        _field_html("Child Project Name", '<input type="text" name="child_name" placeholder="e.g. Foundation Works">'),
        _field_html("Parent Project",  f'<select name="project_id">{proj_sel_opts}</select>'),
        _field_html("Parent Name (display)", '<input type="text" name="parent_name" placeholder="e.g. Skyway Expressway Phase I">'),
      ])}
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:14px">
        {_field_html("Original Value (₹)", '<input type="number" name="original_value" id="vp-orig" placeholder="0" step="1" oninput="calcRevised()">')}
        {_field_html("Variation Amount (₹)", '<input type="number" name="variation_amt" id="vp-vamt" placeholder="0 or negative" step="1" oninput="calcRevised()">')}
      </div>
      <div style="background:#f0f4ff;border:1px solid #c8d8f8;border-radius:8px;padding:12px 16px;margin:14px 0">
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']};margin-bottom:5px">Revised Value (auto)</div>
        <div style="font-size:20px;font-weight:800;color:{C['dark']}" id="vp-revised-val">₹ 0.00</div>
      </div>
      {_field_html("Date of Variation", f'<input type="date" name="variation_date" value="{date.today().isoformat()}">')}
    </form>
  </div>
  <div style="padding:16px 20px;border-top:1px solid {C['outline']};background:{C['surface_low']};display:flex;gap:10px">
    <button onclick="closeVarPanel()" style="flex:1;padding:10px;border:1px solid {C['outline']};border-radius:6px;background:#fff;color:{C['on_var']};font-size:11px;font-weight:700;text-transform:uppercase;cursor:pointer">CANCEL</button>
    <button onclick="document.getElementById('var-form').submit()" style="flex:1;padding:10px;background:{C['dark']};color:#fff;border:none;border-radius:6px;font-size:11px;font-weight:700;text-transform:uppercase;cursor:pointer">SAVE VARIATION</button>
  </div>
</div>
<!-- ── CHILD DETAIL MODAL ── -->
<div id="child-modal-overlay"
     onclick="closeChildModal()"
     style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);
            z-index:500;backdrop-filter:blur(2px);align-items:center;justify-content:center"></div>

<div id="child-modal"
     style="display:none;position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);
            width:min(92vw,900px);max-height:85vh;background:#fff;border-radius:14px;
            box-shadow:0 24px 80px rgba(0,0,0,.22);z-index:501;flex-direction:column;overflow:hidden">

  <!-- header -->
  <div style="padding:18px 24px;border-bottom:1px solid {C['outline']};
              background:{C['surface_low']};display:flex;align-items:center;justify-content:space-between;flex-shrink:0">
    <div>
      <p style="font-size:11px;font-weight:700;text-transform:uppercase;
                letter-spacing:.07em;color:{C['on_var']};margin-bottom:2px">Parent Project</p>
      <h3 id="modal-parent-title"
          style="font-size:18px;font-weight:800;color:{C['dark']}">—</h3>
      <div id="modal-parent-summary"
           style="display:flex;flex-wrap:wrap;gap:16px;margin-top:10px;font-size:12px;color:#475569">
        <span style="font-weight:700;color:{C['dark']}">Order Value: ₹—</span>
        <span style="font-weight:700;color:{C['dark']}">Total Expenditure Billed: ₹—</span>
        <span style="font-weight:700;color:{C['dark']}">Collection: ₹—</span>
      </div>
    </div>
    <button onclick="closeChildModal()"
            style="width:32px;height:32px;border-radius:50%;border:1px solid {C['outline']};
                   background:#fff;font-size:18px;cursor:pointer;color:{C['on_var']};
                   display:flex;align-items:center;justify-content:center">✕</button>
  </div>

  <!-- summary bar -->
  <div id="modal-summary-bar"
       style="display:flex;gap:0;border-bottom:1px solid {C['outline']};flex-shrink:0;overflow-x:auto"></div>

  <!-- child cards scroll area -->
  <div id="modal-body"
       style="overflow-y:auto;padding:20px 24px;flex:1;display:flex;flex-direction:column;gap:16px"></div>

  <!-- footer -->
  <div style="padding:12px 24px;border-top:1px solid {C['outline']};background:{C['surface_low']};
              font-size:11px;color:{C['on_var']};flex-shrink:0">
    All figures in ₹ &nbsp;·&nbsp; Revised Order Value = Original + Variation &nbsp;·&nbsp; EBF = 1.35 × Total Expenditure Billed
  </div>
</div>
<div class="overlay" id="var-overlay" onclick="closeVarPanel()"></div>
<script>
function filterParentRows(){{
  var qv = (document.getElementById('pp-search').value || '').trim().toLowerCase();
  var pmf = document.getElementById('pp-pm-filter').value;
  var rows = document.querySelectorAll('#pp-tbody tr.parent-fin-row');
  var anyVisible = false;
  rows.forEach(function(r){{
    var hay = r.getAttribute('data-search') || '';
    var rowPm = r.getAttribute('data-pm') || '';
    var matchText = !qv || hay.indexOf(qv) !== -1;
    var matchPm   = !pmf || rowPm === pmf;
    var match = matchText && matchPm;
    r.style.display = match ? '' : 'none';
    if (match) anyVisible = true;
    if (!match) {{
      var pid = r.getAttribute('data-pid');
      document.querySelectorAll('#pp-tbody tr.child-of-' + pid).forEach(function(k){{
        k.style.display = 'none';
      }});
    }}
  }});
  document.getElementById('pp-no-match').style.display = anyVisible ? 'none' : '';
}}
function openVarPanel(){{document.getElementById('var-panel').classList.add('open');document.getElementById('var-overlay').classList.add('show');}}
function closeVarPanel(){{document.getElementById('var-panel').classList.remove('open');document.getElementById('var-overlay').classList.remove('show');}}
document.addEventListener('keydown',function(e){{if(e.key==='Escape')closeVarPanel();}});
function calcRevised(){{
  var o=parseFloat(document.getElementById('vp-orig').value)||0;
  var v=parseFloat(document.getElementById('vp-vamt').value)||0;
  document.getElementById('vp-revised-val').textContent='₹ '+(o+v).toLocaleString('en-IN',{{minimumFractionDigits:2}});
}}
var CHILD_DATA = {modal_data_json};

function openChildModal(pid) {{
  var d = CHILD_DATA[pid];
  if (!d || !d.children.length) return;

  document.getElementById('modal-parent-title').textContent = d.parent_co + ': ' + d.parent_name;
  document.getElementById('modal-parent-summary').innerHTML =
    '<span style="font-weight:700;color:#1f2937">Order Value: ₹'+fmt(d.order_value)+'</span>' +
    '<span style="font-weight:700;color:#1f2937">Total Expenditure Billed: ₹'+fmt(d.itd_billed)+'</span>' +
    '<span style="font-weight:700;color:#1f2937">Collection: ₹'+fmt(d.collection)+'</span>';

  // ── summary bar ──
  var totalOV  = d.children.reduce(function(s,c){{return s+c.revised_ov;}}, 0);
  var totalITD = d.children.reduce(function(s,c){{return s+c.billed;}}, 0);
  var totalCol = d.children.reduce(function(s,c){{return s+c.collection;}}, 0);
  var totalVar = d.children.reduce(function(s,c){{return s+c.variation;}}, 0);
  var bar = document.getElementById('modal-summary-bar');
  bar.innerHTML = [
    ['Sub-projects',   d.children.length,              '#1a2540', '#fff'],
    ['Total Rev. OV',  '₹'+fmt(totalOV),               '#1e3a8a', '#e8f0fe'],
    ['Net Variation',  (totalVar>=0?'+':'')+'₹'+fmt(Math.abs(totalVar)), totalVar<0?'#166534':'#991b1b', totalVar<0?'#f0fdf4':'#fef2f2'],
    ['Total Expenditure Billed', '₹'+fmt(totalITD),     '#1a2540', '#f8faff'],
    ['Collection',     '₹'+fmt(totalCol),              '#1a2540', '#f8faff'],
  ].map(function(item){{
    return '<div style="flex:1;min-width:140px;padding:12px 18px;border-right:1px solid #e5e7eb;background:'+item[3]+'">'
      + '<div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:#6b7280;margin-bottom:4px">'+item[0]+'</div>'
      + '<div style="font-size:18px;font-weight:800;color:'+item[2]+'">'+item[1]+'</div>'
      + '</div>';
  }}).join('');

  // ── child cards ──
  var body = document.getElementById('modal-body');

  try {{
    body.innerHTML = d.children.map(function(c) {{
    try {{
      var expPct   = c.revised_ov ? (c.itd_cost / c.revised_ov * 100) : 0;
      var barClr   = expPct >= 100 ? '#dc2626' : expPct >= 80 ? '#f97316' : '#166534';
      var stClr    = (c.status==='Cancelled'||c.status==='Hold')
                     ? 'background:#ffdad6;color:#ba1a1a'
                     : (c.status==='Closed'||c.status==='Execution Closed')
                     ? 'background:#e5e7eb;color:#374151'
                     : 'background:#e0f2fe;color:#0369a1';
      var varClr   = c.variation < 0 ? '#166534' : c.variation > 0 ? '#991b1b' : '#6b7280';
      var varSign  = c.variation > 0 ? '+' : '';
      var warnHTML = expPct >= 100
        ? '<span style="background:#fef2f2;color:#991b1b;border:1px solid #fca5a5;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700">🚨 Over Budget</span>'
        : expPct >= 80
        ? '<span style="background:#fff7ed;color:#c2410c;border:1px solid #fdba74;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700">⚠ High Spend</span>'
        : '<span style="background:#f0fdf4;color:#166534;border:1px solid #bbf7d0;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700">✓ On Track</span>';

      return '<div style="border:1px solid #e5e7eb;border-radius:10px;overflow:hidden;flex-shrink:0">'
        // card header
        + '<div style="padding:12px 16px;background:#f8faff;border-bottom:1px solid #e5e7eb;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px">'
        +   '<div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">'
        +     '<span style="font-size:14px;font-weight:800;color:#1a2540">'+c.co_no+': '+c.name+'</span>'
        +     '<span style="padding:2px 8px;border-radius:3px;font-size:10px;font-weight:700;text-transform:uppercase;'+stClr+'">'+c.status+'</span>'
        +     '<span style="font-size:11px;color:#6b7280">PM: '+c.pm+'</span>'
        +     warnHTML
        +   '</div>'
        + '</div>'
        // kpi grid
        + '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:0;border-bottom:1px solid #e5e7eb">'
        + [
            ['Order Value',    '₹'+fmt(c.order_value),  '#1a2540'],
            ['Variation',      varSign+'₹'+fmt(Math.abs(c.variation)), varClr],
            ['Revised OV',     '₹'+fmt(c.revised_ov),   '#1e3a8a'],
            ['Baseline EAC',   '₹'+fmt(c.baseline_eac), '#1a2540'],
            ['Rev. EAC',       '₹'+fmt(c.rev_eac),      '#1a2540'],
            ['Margin %',       Number(c.margin||0).toFixed(1)+'%',  '#166534'],
            ['Total Expense',  '₹'+fmt(c.itd_cost),     '#1a2540'],
            ['Total Expenditure Billed', '₹'+fmt(c.billed), '#1a2540'],
            ['Collection',     '₹'+fmt(c.collection),   '#1a2540'],
            ['EBF (1.35×)',    '₹'+fmt(c.ebf),          '#1a2540'],
            ['Phys. Complete', Number(c.phys_pct||0).toFixed(1)+'%','#1e3a8a'],
          ].map(function(kpi){{
            return '<div style="padding:10px 14px;border-right:1px solid #f0f0f0">'
              + '<div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#6b7280;margin-bottom:3px">'+kpi[0]+'</div>'
              + '<div style="font-size:14px;font-weight:700;color:'+kpi[2]+';font-family:Courier New,monospace">'+kpi[1]+'</div>'
              + '</div>';
          }}).join('')
        + '</div>'
        // progress bar
        + '<div style="padding:10px 16px">'
        +   '<div style="display:flex;justify-content:space-between;margin-bottom:4px">'
        +     '<span style="font-size:10px;font-weight:700;text-transform:uppercase;color:#6b7280">Budget Utilization (ITD vs Revised OV)</span>'
        +     '<span style="font-size:11px;font-weight:700;color:'+barClr+'">'+expPct.toFixed(1)+'%</span>'
        +   '</div>'
        +   '<div style="background:#e5e7eb;border-radius:99px;height:7px;overflow:hidden">'
        +     '<div style="height:100%;width:'+Math.min(expPct,100).toFixed(1)+'%;background:'+barClr+';border-radius:99px;transition:width .4s"></div>'
        +   '</div>'
        + '</div>'
        + '</div>';
    }} catch (err) {{
      // A single malformed sub-project record must never blank out the whole
      // list — show a visible, identifiable error card instead and log details.
      console.error('Failed to render sub-project card', c, err);
      return '<div style="border:1px solid #fca5a5;border-radius:10px;padding:14px 16px;background:#fef2f2;flex-shrink:0">'
        + '<div style="font-size:13px;font-weight:800;color:#991b1b">⚠ Could not render: '
        +   (c && c.co_no ? c.co_no : '(unknown co_no)') + ' — ' + (c && c.name ? c.name : '(unknown name)')
        + '</div>'
        + '<div style="font-size:11px;color:#7f1d1d;margin-top:4px;font-family:monospace">' + (err && err.message ? err.message : err) + '</div>'
        + '</div>';
    }}
  }}).join('');
  }} catch (outerErr) {{
    console.error('openChildModal: top-level failure', outerErr);
    body.innerHTML =
      '<div style="border:1px solid #fca5a5;border-radius:10px;padding:14px 16px;background:#fef2f2;font-family:monospace;font-size:12px;color:#7f1d1d;white-space:pre-wrap">'
      + 'Top-level render error: ' + (outerErr && outerErr.message ? outerErr.message : outerErr)
      + '</div>';
  }}


  document.getElementById('child-modal').style.display        = 'flex';
  document.getElementById('child-modal-overlay').style.display = 'flex';
  document.body.style.overflow = 'hidden';
}}

function closeChildModal() {{
  document.getElementById('child-modal').style.display        = 'none';
  document.getElementById('child-modal-overlay').style.display = 'none';
  document.body.style.overflow = '';
}}

document.addEventListener('keydown', function(e) {{
  if (e.key === 'Escape') {{ closeChildModal(); closeVarPanel(); }}
}});

function fmt(n) {{
  return (n||0).toLocaleString('en-IN', {{minimumFractionDigits:0, maximumFractionDigits:0}});
}}

function toggleFin(row) {{
  // kept as no-op — expand is now handled by modal
}}
</script>
</body></html>"""
    return html


# ═══════════════════════════════════════════════════════════════
#  PAGE 6 — KPI MONITORING
# ═══════════════════════════════════════════════════════════════
import json

KPI_DEFS = [
    ("RF1","EAC < Total Expenditure",          "EAC is lower than Total Expenditure (Inception-to-Date Cost)"),
    ("RF2","Total Expenditure > Revenue",      "Total Expenditure exceeds up-to-date Revenue Collection"),
    ("RF3","Total Expenditure > Order Val",    "Total Expenditure has surpassed Order Value"),
    ("RF4","%age Physical Completion below %Financial Expenditure", "%age Physical Completion below %Financial Expenditure"),
]


def _fmt_amt(v: float) -> str:
    v = _n(v)
    if abs(v) >= 1e7:
        return f"₹{v/1e7:,.2f}Cr"
    return f"₹{v:,.0f}"


def _eval_kpis(proj: dict, ebf_mult: float = None):
    """Returns (flags, reasons) — flags is a list[bool], reasons is a dict
    mapping RF-code -> human readable explanation, but ONLY for codes that
    are flagged (True). This powers the click-to-explain popup.

    ebf_mult: the EBF multiplier to use for this project. If not supplied,
    it's looked up from the project's own LOB (falls back to 1.35 if that
    LOB has no custom value set via Financials → Project Performance Analysis)."""
    if ebf_mult is None:
        ebf_mult = get_ebf_mult(proj.get("hier_business_line") or proj.get("business_line") or "")
    ov      = _n(proj.get("order_value"))
    rev_ov  = _n(proj.get("revised_order_value")) or (ov + _n(proj.get("variation_amt")))
    itd     = _n(proj.get("itd_cost"))                          # Col Z — Total ITD Cost
    col     = _n(proj.get("collection_till_date"))              # Col S
    eac     = _n(proj.get("rev_latest_eac") or proj.get("baseline_eac"))   # Col N
    phys    = _n(proj.get("physical_completion"))               # Col AE (e.g. 0.60 fraction or 60.0 pct)

    # Normalise physical completion to 0-1 fraction
    phys_frac = phys / 100.0 if phys > 1 else phys
    # Keep a % version for RF4 comparisons
    phys_pct  = phys_frac * 100.0

    # Col AD — Financial Completion %
    fin_completion_pct = (itd / eac * 100) if eac else 0

    # RF4 threshold: physical completion should be >= 90% of financial completion
    phys_thresh = fin_completion_pct * 0.9

    flags = [
        bool(eac > 0 and eac < itd),            # RF1: EAC < ITD Cost
        bool(itd > col),                          # RF2: ITD Cost > Collection
        bool(itd > rev_ov),                       # RF3: ITD Cost > Revised OV
        bool(phys_pct < phys_thresh),             # RF4: Phys% < 90% of Financial%
    ]

    reasons = {}
    if flags[0]:
        reasons["RF1"] = (f"EAC of {_fmt_amt(eac)} is lower than the Total Expenditure of "
                           f"{_fmt_amt(itd)} — the project is forecast to cost more than its own "
                           f"current estimate at completion.")
    if flags[1]:
        reasons["RF2"] = (f"Total Expenditure of {_fmt_amt(itd)} exceeds Revenue Collection of {_fmt_amt(col)} "
                           f"— money spent on the project has outpaced money actually collected.")
    if flags[2]:
        reasons["RF3"] = (f"Total Expenditure of {_fmt_amt(itd)} has surpassed the Revised Order Value of "
                           f"{_fmt_amt(rev_ov)} — cumulative cost is now larger than the contract value.")
    if flags[3]:
        reasons["RF4"] = (f"Physical Completion of {phys_pct:.1f}% is below the minimum threshold of "
                           f"{phys_thresh:.1f}% (90% of Financial Completion, which stands at "
                           f"{fin_completion_pct:.1f}%) — work on the ground is lagging the money spent.")
    return flags, reasons


def kpi(alert_html="", username="Admin") -> str:
    # A "pm" login only ever sees their own projects.
    role    = (getattr(_req_ctx, "role", "") or "").lower()
    own_pm  = (getattr(_req_ctx, "pm_name", "") or "").strip()
    is_scoped_pm = bool(role == "pm" and own_pm)

    projects  = fetch_all_projects()
    if is_scoped_pm:
        projects = [pr for pr in projects if (pr.get("project_manager") or "") == own_pm]
    lob_options_html = "".join(f'<option value="{lo}">{lo}</option>' for lo in get_visible_lob_list())
    ebf_mult_map = get_all_ebf_mults()
    def _lob_ebf_mult(pr):
        lob = pr.get("hier_business_line") or pr.get("business_line") or ""
        return ebf_mult_map.get(lob, DEFAULT_EBF_MULT)
    evaluated = [(p, *_eval_kpis(p, _lob_ebf_mult(p))) for p in projects]   # (proj, flags, reasons)
    total     = len(evaluated) or 1
    red_count = sum(1 for _, f, _ in evaluated if any(f))
    risk_pct  = round(red_count / total * 100)
    # Unbilled Exposure — matches Executive Performance Dashboard's "Total
    # Pending Collection": billed-but-not-yet-collected, summed over parent
    # projects only (so sub-project figures already rolled into a parent
    # aren't double-counted), then floored at 0 once on the aggregate.
    _kpi_parents = _rep_parents_only(projects)
    _kpi_billed  = sum(_n(p.get("billed_amount"))        for p in _kpi_parents)
    _kpi_col     = sum(_n(p.get("collection_till_date")) for p in _kpi_parents)
    unbilled  = max(_kpi_billed - _kpi_col, 0)
    ub_fmt    = f"₹{unbilled/1e7:.2f}Cr" if unbilled >= 1e7 else f"₹{unbilled:,.0f}"
    gauge_col = C['error'] if risk_pct >= 60 else C['secondary'] if risk_pct >= 30 else "#166534"
    gauge_off = round(125.66 * (1 - risk_pct / 100), 2)

    rf_heads = "".join(
        f'<th style="text-align:center;min-width:52px;cursor:help;position:relative">'
        f'{code}'
        f'<div style="display:none;position:absolute;top:calc(100%+4px);left:50%;transform:translateX(-50%);background:#2d3133;color:#eff1f3;padding:6px 10px;border-radius:5px;font-size:10px;width:140px;z-index:40;white-space:normal;line-height:1.4;box-shadow:0 2px 8px rgba(0,0,0,.3)" class="rf-tip">{tip}</div>'
        f'</th>'
        for code, _, tip in KPI_DEFS
    )

    tbody = ""
    reasons_map = {}   # pid -> {code: reason}  — fed into JS for the popup
    for idx, (proj, flags, reasons) in enumerate(evaluated):
        pid  = str(idx)
        co   = proj.get("co_no","—"); name = proj.get("project_name","—")
        lob  = proj.get("business_line") or proj.get("hier_business_line") or "—"
        if reasons:
            reasons_map[pid] = {"name": name, "co": co, "items": reasons}

        cells = "".join(
            (
                f'<td style="text-align:center">'
                f'<span onclick="showFlag(\'{pid}\',\'{KPI_DEFS[i][0]}\')" '
                f'style="font-size:18px;color:{C["error"]};opacity:1;cursor:pointer" '
                f'title="Click to see why">⚑</span></td>'
                if f else
                f'<td style="text-align:center">'
                f'<span style="font-size:16px;color:#16a34a;opacity:.85">✓</span></td>'
            )
            for i, f in enumerate(flags)
        )
        pm_name = proj.get("project_manager") or ""
        tbody += f"""
        <tr data-lob="{lob}" data-pm="{pm_name}">
          <td style="position:sticky;left:0;background:#fff;padding:12px 14px 12px 20px;min-width:220px;border-bottom:1px solid #f0f2f4">
            <div style="font-weight:700;color:{C['dark']};font-size:14px">{name}</div>
            <div style="font-size:12px;color:{C['on_var']}">{lob}</div>
          </td>
          <td style="font-family:'Courier New',monospace;font-size:13px;color:{C['on_var']};padding:12px 10px;border-bottom:1px solid #f0f2f4">{co}</td>
          {cells}
          <td style="text-align:right;padding:12px 16px;border-bottom:1px solid #f0f2f4">
            <span style="font-size:18px;cursor:pointer;color:{C['on_var']}">⊞</span>
          </td>
        </tr>"""

    if not tbody:
        tbody = f'<tr><td colspan="7" style="padding:24px;text-align:center;color:#aaa">No projects yet. <a href="/data_entry">Add projects →</a></td></tr>'

    half       = len(KPI_DEFS) // 2
    meth_left  = "".join(
        f'<div style="display:flex;align-items:flex-start;gap:10px;margin-bottom:10px">'
        f'<span style="background:{C["dark"]};color:#fff;font-size:10px;font-weight:700;padding:2px 6px;border-radius:4px;flex-shrink:0;font-family:monospace">{code}</span>'
        f'<p style="font-size:13px;color:{C["on_surface"]};line-height:1.5">{desc}</p></div>'
        for code, _, desc in KPI_DEFS[:half]
    )
    meth_right = "".join(
        f'<div style="display:flex;align-items:flex-start;gap:10px;margin-bottom:10px">'
        f'<span style="background:{C["dark"]};color:#fff;font-size:10px;font-weight:700;padding:2px 6px;border-radius:4px;flex-shrink:0;font-family:monospace">{code}</span>'
        f'<p style="font-size:13px;color:{C["on_surface"]};line-height:1.5">{desc}</p></div>'
        for code, _, desc in KPI_DEFS[half:]
    )

    # Lookup of RF code -> short label, used by the popup to title itself
    rf_labels = {code: tip for code, _, tip in KPI_DEFS}

    html  = _html_head("KPI Monitor") + "<body>"
    html += _sidebar("kpi") + _topbar("KPI MONITORING", username)
    html += f"""
<style>
.kpi-bento{{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:24px}}
.kpi-card{{background:#fff;border:1px solid {C['outline']};border-radius:10px;padding:20px 22px}}
.kpi-tbl{{width:100%;border-collapse:collapse;font-size:13px}}
.kpi-tbl thead tr{{background:{C['surface_low']};border-bottom:1px solid {C['outline']}}}
.kpi-tbl th{{padding:11px 10px;font-size:10px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['on_var']};white-space:nowrap}}
th:hover .rf-tip{{display:block!important}}
#flag-modal-backdrop{{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:99;display:none;align-items:center;justify-content:center}}
#flag-modal{{background:#fff;border-radius:12px;max-width:420px;width:90%;padding:22px 24px;box-shadow:0 12px 40px rgba(0,0,0,.25)}}
#flag-modal .fm-code{{display:inline-block;background:{C['error']};color:#fff;font-size:11px;font-weight:700;padding:3px 8px;border-radius:5px;font-family:monospace;margin-bottom:10px}}
#flag-modal h4{{font-size:16px;font-weight:700;color:{C['dark']};margin-bottom:2px}}
#flag-modal .fm-sub{{font-size:12px;color:{C['on_var']};margin-bottom:14px}}
#flag-modal p.fm-reason{{font-size:14px;line-height:1.6;color:{C['on_surface']};background:{C['surface_low']};padding:12px 14px;border-radius:8px;border-left:3px solid {C['error']}}}
#flag-modal .fm-close{{margin-top:18px;width:100%;padding:9px;border-radius:7px;border:1px solid {C['outline']};background:#fff;font-size:13px;font-weight:600;cursor:pointer;color:{C['dark']}}}
</style>
<div class="main"><div class="canvas" style="max-width:1360px">
  {alert_html}
  <div style="display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:20px">
    <div>
      <h2 style="font-size:28px;font-weight:700;color:{C['dark']}">KPI Monitor</h2>
      <p style="font-size:14px;color:#6b7280;margin-top:4px">Red flag tracking across {total} live projects from DB</p>
    </div>
  </div>

  <!-- Filter bar: search + LOB -->
  <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;background:#fff;
              border:1px solid {C['outline']};border-radius:10px;padding:12px 16px;margin-bottom:20px">
    <input id="kpi-search" type="text" placeholder="🔎 Search by project name or CO.No…"
           oninput="filterKpiTable()"
           style="flex:1;min-width:220px;padding:8px 12px;border:1px solid {C['outline']};
                  border-radius:7px;font-size:13px;outline:none;background:{C['surface_low']};color:{C['dark']}">
    <div style="width:1px;height:24px;background:{C['outline']}"></div>
    <span style="font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['on_var']}">LOB</span>
    <select id="kpi-lob-filter" onchange="filterKpiTable()"
            style="padding:8px 12px;border:1px solid {C['outline']};border-radius:7px;font-size:12px;
                   background:{C['surface_low']};color:{C['dark']};outline:none;cursor:pointer;min-width:170px">
      <option value="">All LOB</option>
      {lob_options_html}
    </select>
    <div style="width:1px;height:24px;background:{C['outline']}"></div>
    <span style="font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{C['on_var']}">PM</span>
    {f'''<div style="padding:8px 12px;border:1px solid {C['outline']};border-radius:7px;font-size:12px;
                background:{C['surface_low']};color:{C['dark']};font-weight:600;display:inline-flex;
                align-items:center;gap:6px">👤 {own_pm}</div>''' if is_scoped_pm else f'''<select id="kpi-pm-filter" onchange="filterKpiTable()"
            style="padding:8px 12px;border:1px solid {C['outline']};border-radius:7px;font-size:12px;
                   background:{C['surface_low']};color:{C['dark']};outline:none;cursor:pointer;min-width:170px">
      <option value="">All PMs</option>
      {"".join(f'<option value="{pmn}">{pmn}</option>' for pmn in get_project_managers())}
    </select>'''}
  </div>

  <div class="kpi-bento">
    <div class="kpi-card">
      <p style="font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['on_var']};margin-bottom:10px">Total Red Flag Projects</p>
      <div style="display:flex;align-items:flex-end;gap:10px">
        <span style="font-size:32px;font-weight:700;line-height:1;color:{C['error']}">{red_count}</span>
        <span style="font-size:13px;color:{C['on_var']};margin-bottom:4px">of {total} projects</span>
      </div>
      <div style="margin-top:14px;padding-top:12px;border-top:1px solid {C['outline']};font-size:13px;font-weight:700;color:{C['error']}">↑ {risk_pct}% flagged</div>
    </div>
    <div class="kpi-card" style="display:flex;flex-direction:column;align-items:center">
      <p style="font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['on_var']};margin-bottom:10px;align-self:flex-start">Critical Risk Level</p>
      <div style="position:relative">
        <svg viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg" style="width:180px;height:90px">
          <path fill="none" stroke="{C['surface_high']}" stroke-width="12" stroke-linecap="round" d="M 10 50 A 40 40 0 0 1 90 50"/>
          <path fill="none" stroke="{gauge_col}" stroke-width="12" stroke-linecap="round"
                stroke-dasharray="125.66" stroke-dashoffset="125.66"
                d="M 10 50 A 40 40 0 0 1 90 50"
                style="transition:stroke-dashoffset 1s ease-in-out" id="gauge-arc"/>
        </svg>
        <div style="position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;padding-bottom:4px">
          <span style="font-size:24px;font-weight:700;color:{gauge_col};line-height:1">{risk_pct}%</span>
          <span style="font-size:11px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">{("High Risk" if risk_pct>=60 else "Medium" if risk_pct>=30 else "Low Risk")}</span>
        </div>
      </div>
    </div>
    <div class="kpi-card">
      <p style="font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['on_var']};margin-bottom:10px">Unbilled Exposure</p>
      <div style="display:flex;align-items:flex-end;gap:10px">
        <span style="font-size:32px;font-weight:700;line-height:1;color:{C['dark']}">{ub_fmt}</span>
        <span style="font-size:13px;color:{C['on_var']};margin-bottom:4px">Current Outstandings</span>
      </div>
    </div>
  </div>

  <div style="background:#fff;border:1px solid {C['outline']};border-radius:8px;overflow:hidden;margin-bottom:24px">
    <div style="padding:14px 20px;border-bottom:1px solid {C['outline']};display:flex;align-items:center;justify-content:space-between;background:{C['surface_low']}">
      <h3 style="font-size:20px;font-weight:700;color:{C['dark']}">KPI Red Flag Monitor</h3>
      <a href="/export_excel/kpi"
         style="display:inline-flex;align-items:center;gap:6px;padding:7px 14px;
                background:{C['dark']};color:#fff;border-radius:7px;font-size:11px;
                font-weight:700;text-decoration:none;letter-spacing:.04em"
      >⬇ Export Excel</a>
    </div>
    <div style="overflow-x:auto">
      <table class="kpi-tbl">
        <thead><tr>
          <th style="position:sticky;left:0;background:{C['surface_low']};z-index:20;min-width:240px;text-align:left;padding-left:20px">Project Name</th>
          <th style="min-width:90px;text-align:left;padding-left:14px">CO.NO.</th>
          {rf_heads}
          <th style="text-align:right;padding-right:16px">Actions</th>
        </tr></thead>
        <tbody id="kpi-tbody">{tbody}</tbody>
      </table>
    </div>
    <div style="padding:12px 20px;background:{C['surface_low']};border-top:1px solid {C['outline']};display:flex;justify-content:space-between;align-items:center">
      <span style="font-size:12px;color:{C['on_var']}">Showing <span id="kpi-visible-count">{total}</span> of {total} projects · {red_count} red flags</span>
    </div>
  </div>

  <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
    <div style="background:{C['surface_low']};border:1px solid {C['outline']};border-radius:10px;padding:18px">
      <h4 style="font-size:11px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['dark']};margin-bottom:14px">Red Flag Methodology</h4>
      {meth_left}
    </div>
    <div style="background:{C['surface_low']};border:1px solid {C['outline']};border-radius:10px;padding:18px">
      <h4 style="font-size:11px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['dark']};margin-bottom:14px;visibility:hidden">Methodology Continued</h4>
      {meth_right}
    </div>
  </div>
</div></div>

<div id="flag-modal-backdrop" onclick="if(event.target===this)closeFlagModal()">
  <div id="flag-modal">
    <span class="fm-code" id="fm-code">RF1</span>
    <h4 id="fm-title">Project Name</h4>
    <div class="fm-sub" id="fm-sub">CO.NO. —</div>
    <p class="fm-reason" id="fm-reason">Reason goes here.</p>
    <button class="fm-close" onclick="closeFlagModal()">Close</button>
  </div>
</div>

<script>
window.addEventListener('load',function(){{
  var arc=document.getElementById('gauge-arc');
  if(arc){{arc.style.strokeDashoffset='125.66';setTimeout(function(){{arc.style.strokeDashoffset='{gauge_off}';}},400);}}
}});
function filterKpiTable(){{
  var q   = (document.getElementById('kpi-search').value || '').toLowerCase().trim();
  var lob = document.getElementById('kpi-lob-filter').value;
  var pmEl = document.getElementById('kpi-pm-filter');
  var pmf = pmEl ? pmEl.value : '';
  var visible = 0;
  document.querySelectorAll('#kpi-tbody tr').forEach(function(r){{
    var matchesText = !q || r.textContent.toLowerCase().includes(q);
    var matchesLob  = !lob || r.getAttribute('data-lob') === lob;
    var matchesPm   = !pmf || r.getAttribute('data-pm') === pmf;
    var show = matchesText && matchesLob && matchesPm;
    r.style.display = show ? '' : 'none';
    if (show) visible++;
  }});
  var countEl = document.getElementById('kpi-visible-count');
  if (countEl) countEl.textContent = visible;
}}

const KPI_REASONS = {json.dumps(reasons_map, ensure_ascii=False)};
const RF_LABELS = {json.dumps(rf_labels, ensure_ascii=False)};

function showFlag(pid, code){{
  const proj = KPI_REASONS[pid];
  if(!proj || !proj.items[code]) return;
  document.getElementById('fm-code').textContent = code;
  document.getElementById('fm-title').textContent = proj.name;
  document.getElementById('fm-sub').textContent = 'CO.NO. ' + proj.co + ' · ' + (RF_LABELS[code] || '');
  document.getElementById('fm-reason').textContent = proj.items[code];
  document.getElementById('flag-modal-backdrop').style.display = 'flex';
}}
function closeFlagModal(){{
  document.getElementById('flag-modal-backdrop').style.display = 'none';
}}
document.addEventListener('keydown', function(e){{
  if(e.key === 'Escape') closeFlagModal();
}});

(function(){{
  var PAGE = 'kpi';
  var saved = MonarchFilters.init(PAGE);
  if (saved) {{
    var s = document.getElementById('kpi-search');
    var l = document.getElementById('kpi-lob-filter');
    if (s && saved.search) s.value = saved.search;
    if (l && saved.lob)    l.value = saved.lob;
    filterKpiTable();
  }}
  function saveKpi() {{
    MonarchFilters.save(PAGE, {{
      search: (document.getElementById('kpi-search')||{{}}).value || '',
      lob:    (document.getElementById('kpi-lob-filter')||{{}}).value || ''
    }});
  }}
  var s2 = document.getElementById('kpi-search');
  var l2 = document.getElementById('kpi-lob-filter');
  if (s2) s2.addEventListener('input', saveKpi);
  if (l2) l2.addEventListener('change', saveKpi);
}})();
</script>
</body></html>"""
    return html
# ═══════════════════════════════════════════════════════════════
""" page 7 REPORTS"""
# ═══════════════════════════════════════════════════════════════

import io
import json as _json
from openpyxl import Workbook as _WB
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── Excel export (replaces PDF export) ───────────────────────────────────────

# ── Shared helpers for Consolidated Reports (used by page + export) ─────────
def _rep_oi_amt(p):
    rev = _n(p.get("revised_order_value"))
    return rev if rev else _n(p.get("order_value"))

def _rep_lob(p):
    return (p.get("hier_business_line") or p.get("business_line") or "").strip()

def _rep_phy_pct(p):
    v = _n(p.get("physical_completion"))
    v = v * 100 if v <= 1 else v          # normalise fraction (0.6) -> percent (60)
    return round(v, 1)

def _rep_fin_pct(p):
    itd_cost = _n(p.get("itd_cost"))
    eac      = _n(p.get("rev_latest_eac") or p.get("baseline_eac"))
    if eac:
        return round(itd_cost / eac * 100, 1)
    ov       = _rep_oi_amt(p) or _rev_ov(p)
    return round((itd_cost / ov * 100) if ov else 0, 1)

def _rep_duration_days(p):
    end_str = p.get("extended_end_date") or p.get("end_date") or ""
    if not end_str:
        return None
    try:
        end_dt = datetime.fromisoformat(str(end_str)).date()
    except Exception:
        try:
            end_dt = datetime.strptime(str(end_str)[:10], "%Y-%m-%d").date()
        except Exception:
            return None
    return (end_dt - datetime.now().date()).days

def _rep_parents_only(projects):
    return [p for p in projects if p.get("project_type", "parent") != "child"]

def _rep_child_projects(projects):
    """All projects that are linked as a sub-project, by either linkage style used
    across the app: the newer parent_id (co_no) link, or the legacy parent_project
    (project name) text link."""
    return [p for p in projects
            if p.get("project_type") == "child" or (p.get("parent_project") or "").strip()]

def _rep_kids_of(parent, children):
    """Sub-projects belonging to a given parent, matched the same way as the
    Parent Project Details page: by co_no (parent_id) first, falling back to a
    case-insensitive match on the legacy parent_project name field."""
    par_co   = (parent.get("co_no") or "").strip()
    par_name = (parent.get("project_name") or "").strip().lower()
    return [c for c in children
            if (par_co and (c.get("parent_id") or "").strip() == par_co)
            or (c.get("parent_project") or "").strip().lower() == par_name]

def _rep_row(p):
    """Build one report row dict (used for both parent and child project rows)."""
    return {
        "co_no":  p.get("co_no") or "—",
        "name":   p.get("project_name") or "—",
        "lob":    _rep_lob(p),
        "client": p.get("client_name") or "—",
        "date":   (p.get("created_at") or p.get("order_issue_date") or "")[:10] or "—",
        "oi":     _rep_oi_amt(p),
        "col":    _n(p.get("collection_till_date")),
        "itd":    _n(p.get("billed_amount")),
        "bnc":    _n(p.get("billed_amount")) - _n(p.get("collection_till_date")),
        "ub":     max(_rep_oi_amt(p) - _n(p.get("billed_amount")), 0),
        "fin":    _rep_fin_pct(p),
        "phy":    _rep_phy_pct(p),
        "dur":    _rep_duration_days(p),
    }

# key -> (column header, footer label, value-type, value-getter)
REPORT_TAB_ORDER = ["oi", "col", "itd", "bnc", "ub", "fin", "phy", "dur"]
REPORT_TAB_META  = {
    "oi":  ("Order Intake (₹)",    "Total Order Intake",   "currency", _rep_oi_amt),
    "col": ("Collection (₹)",      "Total Collection",     "currency", lambda p: _n(p.get("collection_till_date"))),
    "itd": ("Billed Amount (₹)",   "Total Billed Amount",  "currency", lambda p: _n(p.get("billed_amount"))),
    "bnc": ("Billed but Not Collected (₹)", "Total Billed but Not Collected", "currency", lambda p: _n(p.get("billed_amount")) - _n(p.get("collection_till_date"))),
    "ub":  ("Unbilled Amount (₹)", "Total Unbilled Amount","currency", lambda p: max(_rep_oi_amt(p) - _n(p.get("billed_amount")), 0)),
    "fin": ("% Expenditure Over Order Value", "Avg. % Expenditure Over Order Value", "percent",  _rep_fin_pct),
    "phy": ("Physical %",          "Avg. Physical %",      "percent",  _rep_phy_pct),
    "dur": ("Duration (Days)",     "Avg. Duration (Days)", "days",     _rep_duration_days),
}

# ═══════════════════════════════════════════════════════════════
#  PAGE — SCHEDULE PERFORMANCE REPORT (next to KPI Monitoring)
# ═══════════════════════════════════════════════════════════════

def _sched_parse_date(s):
    s = (s or "").strip()
    if not s:
        return None
    try:
        return datetime.fromisoformat(s[:10]).date()
    except Exception:
        try:
            return datetime.strptime(s[:10], "%Y-%m-%d").date()
        except Exception:
            return None

def _sched_months_between(start_str, end_str):
    d1 = _sched_parse_date(start_str)
    d2 = _sched_parse_date(end_str)
    if not d1 or not d2:
        return None
    months = (d2.year - d1.year) * 12 + (d2.month - d1.month)
    if d2.day < d1.day:
        months -= 1
    return max(months, 0)

def _sched_row_metrics(p) -> dict:
    """Compute WO start/period/extension/% time-elapsed (or % time-overrun)
    metrics for one project, per the scheduling performance formulas:

      wo.start   = Order Issue Date            (data entry)
      wo.period  = Period in Months             (data entry)
      Ext. durn. = Extension Granted, in Months (data entry)

      1) Project End Dt.   = wo.start + wo.period*30 + Ext.durn*30
      2) If Current Date < Project End Dt.:
           % Time Elapsed  = (Current Date - wo.start) / (wo.period*30 + Ext.durn*30)
      3) If Current Date > Project End Dt.:
           % Time Overrun  = (Current Date - Project End Dt.) / (wo.period*30 + Ext.durn*30)
           (flagged / shown in red)
    """
    wo_start = _sched_parse_date(p.get("order_issue_date"))
    end_d    = _sched_parse_date(p.get("end_date"))
    today    = datetime.now().date()

    # wo.period — Period in Months, as captured at data entry
    # (end_date is auto-derived on entry as order_issue_date + period months,
    #  so the entered period is recovered from the two stored dates).
    wo_period = _sched_months_between(p.get("order_issue_date"), p.get("end_date"))

    # Ext. durn. — Extension Granted, in Months (direct data-entry field)
    ext_durn = _n(p.get("extension_months")) or 0

    proj_end_dt  = None
    total_days   = None
    pct_elapsed  = None
    pct_overrun  = None
    is_overrun   = False

    if wo_start is not None and wo_period is not None:
        total_days = wo_period * 30 + ext_durn * 30          # wo.period*30 + Ext.durn*30
        proj_end_dt = wo_start + timedelta(days=total_days)  # 1) Project End Dt.

        if total_days > 0:
            if today < proj_end_dt:
                # 2) % Time Elapsed
                pct_elapsed = round((today - wo_start).days / total_days * 100, 1)
            elif today > proj_end_dt:
                # 3) % Time Overrun (red)
                pct_elapsed = 100.0
                pct_overrun = round((today - proj_end_dt).days / total_days * 100, 1)
                is_overrun  = True
            else:
                pct_elapsed = 100.0

    ext_days = round(ext_durn * 30)

    phy_pct    = _rep_phy_pct(p)
    is_overdue = is_overrun and phy_pct < 100

    return {
        "start_d": wo_start, "end_d": end_d, "ext_d": proj_end_dt,
        "period_months": wo_period, "ext_months": ext_durn, "ext_days": ext_days,
        "revised_end": proj_end_dt, "proj_end_dt": proj_end_dt,
        "pct_elapsed": pct_elapsed, "pct_overrun": pct_overrun, "is_overrun": is_overrun,
        "phy_pct": phy_pct, "is_overdue": is_overdue,
    }

def _sched_filtered_projects(lob="", pm="", status="", q="") -> list:
    projects = _rep_parents_only(fetch_all_projects())
    if lob:
        projects = [p for p in projects if _rep_lob(p) == lob]
    if pm:
        projects = [p for p in projects if (p.get("project_manager") or "") == pm]
    if status:
        projects = [p for p in projects if (p.get("project_status") or "") == status]
    if q:
        qlow = q.strip().lower()
        def _match(p):
            hay = " ".join(str(p.get(k, "") or "") for k in
                            ("id", "co_no", "project_name")).lower()
            return qlow in hay
        projects = [p for p in projects if _match(p)]
    return projects

def export_schedule_excel_bytes(lob="", pm="", status="", q="") -> bytes:
    """Export the Schedule Performance Report — WO dates, extensions,
    physical completion and % time elapsed — straight from the DB."""
    from openpyxl.styles import Font, PatternFill, Alignment
    # A "pm" login only ever sees their own projects — ignore any other
    # pm value that may have been passed (e.g. a hand-edited URL).
    role   = (getattr(_req_ctx, "role", "") or "").lower()
    own_pm = (getattr(_req_ctx, "pm_name", "") or "").strip()
    if role == "pm" and own_pm:
        pm = own_pm
    projects = _sched_filtered_projects(lob, pm, status, q)
    wb = Workbook(); ws = wb.active; ws.title = "Schedule Performance"
    hdr_fill = PatternFill("solid", fgColor="000E24")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    alt_fill = PatternFill("solid", fgColor="F2F4F6")
    headers = ["Project No.", "Project Name", "LOB", "Project Manager", "Status",
               "WO Start", "WO Period (Months)", "Ext. Granted (Months)",
               "Project End Dt.", "Phy. Compln. %", "% Time Elapsed", "% Time Overrun"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = hdr_fill; cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    red_font = Font(color="BA1A1A", bold=True)
    for i, p in enumerate(projects, start=2):
        m = _sched_row_metrics(p)
        ws.append([
            p.get("co_no",""), p.get("project_name",""), _rep_lob(p),
            p.get("project_manager",""), p.get("project_status",""),
            m["start_d"].isoformat() if m["start_d"] else "",
            m["period_months"] if m["period_months"] is not None else "",
            m["ext_months"] if m["ext_months"] else 0,
            m["proj_end_dt"].isoformat() if m["proj_end_dt"] else "",
            m["phy_pct"],
            m["pct_elapsed"] if m["pct_elapsed"] is not None else "",
            m["pct_overrun"] if m["pct_overrun"] is not None else "",
        ])
        if i % 2 == 0:
            for cell in ws[i]: cell.fill = alt_fill
        if m["is_overrun"]:
            ws.cell(i, 12).font = red_font
    for idx, w in enumerate([14, 34, 18, 20, 14, 13, 16, 18, 16, 14, 14, 15], start=1):
        ws.column_dimensions[ws.cell(1, idx).column_letter].width = w
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf.read()

def schedule_performance_report(username="Admin", lob="", pm="", status="", q="", page: int = 1) -> str:
    # A "pm" login only ever sees their own projects — ignore any other
    # pm= value that may have been passed (e.g. a hand-edited URL), and
    # don't offer the other-PM options in the selector at all.
    role   = (getattr(_req_ctx, "role", "") or "").lower()
    own_pm = (getattr(_req_ctx, "pm_name", "") or "").strip()
    is_scoped_pm = bool(role == "pm" and own_pm)
    if is_scoped_pm:
        pm = own_pm

    all_filtered = _sched_filtered_projects(lob, pm, status, q)

    # "Active Projects" and "total portfolio" are meant to be stable, global
    # figures -- they should reflect every project (parents AND children) and
    # must NOT move just because someone applies a LOB/PM/Status/search filter
    # on this page. fetch_all_projects() already applies the one scoping rule
    # that's not a "filter" (a PM login only ever sees their own projects) --
    # everything else (lob, pm, status, q, and the parent/child split used for
    # the table rows below) is deliberately left out of this calculation.
    base_projects = fetch_all_projects()
    total_all     = len(base_projects)
    active_count  = sum(1 for p in base_projects if (p.get("project_status") or "").lower() == "active")

    rows_data = []
    for p in all_filtered:
        m = _sched_row_metrics(p)
        rows_data.append((p, m))

    avg_phy        = (sum(m["phy_pct"] for _, m in rows_data) / len(rows_data)) if rows_data else 0
    overdue_items  = sum(1 for _, m in rows_data if m["is_overdue"])
    on_time_pct    = round((len(rows_data) - overdue_items) / len(rows_data) * 100) if rows_data else 0

    # ── Pagination ──
    PER_PAGE  = 8
    n_rows    = len(rows_data)
    n_pages   = max(1, math.ceil(n_rows / PER_PAGE))
    page      = max(1, min(page, n_pages))
    start_ix  = (page - 1) * PER_PAGE
    page_rows = rows_data[start_ix:start_ix + PER_PAGE]

    def _dfmt(d):
        return d.strftime("%d %b %Y") if d else "—"

    row_html = ""
    for p, m in page_rows:
        co   = p.get("co_no") or "—"
        name = p.get("project_name") or "—"
        lob_name = _rep_lob(p) or "—"

        wo_start  = _dfmt(m["start_d"])
        wo_period = f'{m["period_months"]} Months' if m["period_months"] is not None else "—"

        if m["ext_months"] > 1:
            ext_style = f"color:{C['error']};font-weight:700"
        elif m["ext_months"] > 0:
            ext_style = f"color:{C['secondary']};font-weight:700"
        else:
            ext_style = f"color:{C['on_surface']}"
        ext_str = f'+{m["ext_months"]:g} Months' if m["ext_months"] > 0 else "0 Months"

        ext_dt_style = f"color:{C['error']};font-weight:700" if m["is_overrun"] else f"color:{C['on_surface']}"
        ext_dt_str   = _dfmt(m["proj_end_dt"])

        phy      = m["phy_pct"]
        bar_col  = C["error"] if m["is_overdue"] else C["secondary"]
        phy_bar  = f"""
          <div style="display:flex;align-items:center;gap:8px">
            <div style="width:70px;height:6px;background:{C['surface_high']};border-radius:99px;overflow:hidden">
              <div style="height:100%;width:{min(phy,100):.0f}%;background:{bar_col};border-radius:99px"></div>
            </div>
            <span style="font-size:12px;font-weight:700;color:{C['dark']}">{phy:.0f}%</span>
          </div>"""

        pct_el = m["pct_elapsed"]
        pct_ov = m["pct_overrun"]
        if pct_el is None:
            time_el_html = '<span style="color:#9ca3af">—</span>'
        elif m["is_overrun"]:
            time_el_html = f'<span style="color:{C["error"]};font-weight:700">{pct_ov:.0f}% Overrun</span>'
        else:
            time_el_html = f'<span style="color:{C["on_surface"]}">{pct_el:.0f}%</span>'

        row_html += f"""
        <tr style="border-bottom:1px solid #f0f2f4">
          <td style="padding:14px 16px;font-family:'Courier New',monospace;font-size:12px;color:{C['on_var']}">{co}</td>
          <td style="padding:14px 16px">
            <div style="font-weight:700;color:{C['dark']};font-size:13px">{name}</div>
            <div style="font-size:11px;color:{C['on_var']}">{lob_name}</div>
          </td>
          <td style="padding:14px 16px;font-size:12px;color:{C['on_surface']}">{wo_start}</td>
          <td style="padding:14px 16px;font-size:12px;color:{C['on_surface']}">{wo_period}</td>
          <td style="padding:14px 16px;font-size:12px;{ext_style}">{ext_str}</td>
          <td style="padding:14px 16px;font-size:12px;{ext_dt_style}">{ext_dt_str}</td>
          <td style="padding:14px 16px">{phy_bar}</td>
          <td style="padding:14px 16px;font-size:12px">{time_el_html}</td>
        </tr>"""

    if not row_html:
        row_html = f'<tr><td colspan="8" style="padding:28px;text-align:center;color:#aaa">No projects match these filters.</td></tr>'

    def _qs(pg):
        return (f"?lob={lob}&pm={pm}&status={status}&q={q}&page={pg}").replace(" ", "%20")

    page_links = ""
    window_start = max(1, page - 2)
    window_end   = min(n_pages, page + 2)
    if window_start > 1:
        page_links += f'<a href="/schedule{_qs(1)}" class="sch-pg">1</a><span style="color:#9ca3af">…</span>'
    for pnum in range(window_start, window_end + 1):
        cls = "sch-pg sch-pg-active" if pnum == page else "sch-pg"
        page_links += f'<a href="/schedule{_qs(pnum)}" class="{cls}">{pnum}</a>'
    if window_end < n_pages:
        page_links += f'<span style="color:#9ca3af">…</span><a href="/schedule{_qs(n_pages)}" class="sch-pg">{n_pages}</a>'

    lob_opts    = "".join(f'<option value="{lo}" {"selected" if lo==lob else ""}>{lo}</option>' for lo in get_visible_lob_list())
    pm_opts     = "".join(f'<option value="{p_}" {"selected" if p_==pm else ""}>{p_}</option>' for p_ in get_project_managers())
    status_opts = "".join(f'<option value="{s}" {"selected" if s==status else ""}>{s}</option>' for s in get_project_statuses())

    css = f"""<style>
.sch-kpi{{background:#fff;border:1px solid {C['outline']};border-radius:10px;padding:16px 18px}}
.sch-pg{{display:inline-flex;align-items:center;justify-content:center;min-width:30px;height:30px;
         border-radius:7px;border:1px solid {C['outline']};background:#fff;color:{C['on_var']};
         font-size:12px;font-weight:700;text-decoration:none;padding:0 8px}}
.sch-pg:hover{{background:{C['surface_high']}}}
.sch-pg-active{{background:{C['dark']};color:#fff;border-color:{C['dark']}}}
</style>"""

    html  = _html_head("Schedule Performance Report") + css + "<body>"
    html += _sidebar("schedule") + _topbar("SCHEDULE PERFORMANCE REPORT", username)
    html += f"""
<div class="main"><div class="canvas" style="max-width:1360px">

  <div style="display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:20px">
    <div>
      <h2 style="font-size:26px;font-weight:800;color:{C['dark']}">Schedule Performance Report</h2>
      <p style="font-size:13px;color:#6b7280;margin-top:4px">Track project timelines, extensions, and overdue status across the portfolio.</p>
    </div>
    <a href="/export_excel/schedule?lob={lob}&pm={pm}&status={status}&q={q}"
       style="display:inline-flex;align-items:center;gap:6px;padding:9px 18px;
              background:#fff;border:1px solid {C['outline']};border-radius:8px;font-size:12px;
              font-weight:700;text-decoration:none;color:{C['dark']}">
      ⬇ Export Excel
    </a>
  </div>

  <!-- Filter bar -->
  <form method="get" action="/schedule">
    <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;padding:16px 20px;
                display:flex;align-items:flex-end;gap:16px;flex-wrap:wrap;margin-bottom:20px">
      <div style="flex:1.4;min-width:220px">
        <label style="font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
                      color:{C['on_var']};display:block;margin-bottom:6px">Search (Project ID / CO No. / Name)</label>
        <input type="text" name="q" id="sch-q-input" value="{q.replace('&','&amp;').replace('"','&quot;').replace("'",'&#39;')}" placeholder="e.g. 42, C-24-03, Skyway Expressway"
               style="width:100%;padding:9px 12px;border:1px solid {C['outline']};
                      border-radius:6px;font-size:13px;background:#fff;outline:none;color:{C['dark']}">
      </div>
      <div style="flex:1;min-width:160px">
        <label style="font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
                      color:{C['on_var']};display:block;margin-bottom:6px">Line of Business</label>
        <select name="lob" style="width:100%;padding:9px 12px;border:1px solid {C['outline']};
                border-radius:6px;font-size:13px;background:#fff;outline:none;color:{C['dark']}">
          <option value="">All Businesses</option>{lob_opts}
        </select>
      </div>
      <div style="flex:1;min-width:160px">
        <label style="font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
                      color:{C['on_var']};display:block;margin-bottom:6px">Project Manager</label>
        {f'''<div style="width:100%;padding:9px 12px;border:1px solid {C['outline']};border-radius:6px;
                    font-size:13px;background:{C['surface_low']};color:{C['dark']};font-weight:600;
                    display:flex;align-items:center;gap:6px">👤 {own_pm}</div>
        <input type="hidden" name="pm" value="{own_pm}">''' if is_scoped_pm else f'''<select name="pm" style="width:100%;padding:9px 12px;border:1px solid {C['outline']};
                border-radius:6px;font-size:13px;background:#fff;outline:none;color:{C['dark']}">
          <option value="">Select Manager</option>{pm_opts}
        </select>'''}
      </div>
      <div style="flex:1;min-width:160px">
        <label style="font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
                      color:{C['on_var']};display:block;margin-bottom:6px">Status</label>
        <select name="status" style="width:100%;padding:9px 12px;border:1px solid {C['outline']};
                border-radius:6px;font-size:13px;background:#fff;outline:none;color:{C['dark']}">
          <option value="">All Statuses</option>{status_opts}
        </select>
      </div>
      <button type="submit"
              style="background:{C['secondary']};color:#fff;border:none;border-radius:8px;
                     padding:10px 22px;font-size:12px;font-weight:700;cursor:pointer;
                     display:inline-flex;align-items:center;gap:8px">
        ⚗ Apply Filters
      </button>
    </div>
  </form>

  <!-- KPI cards -->
  <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:22px">
    <div class="sch-kpi">
      <p style="font-size:10px;font-weight:700;letter-spacing:.07em;color:#9ca3af;text-transform:uppercase;margin-bottom:8px">Active Projects</p>
      <p style="font-size:24px;font-weight:800;color:{C['dark']};line-height:1">{active_count}</p>
      <p style="font-size:11px;color:#6b7280;margin-top:6px">of {total_all} total portfolio</p>
    </div>
    <div class="sch-kpi">
      <p style="font-size:10px;font-weight:700;letter-spacing:.07em;color:#9ca3af;text-transform:uppercase;margin-bottom:8px">Avg. Physical Completion</p>
      <p style="font-size:24px;font-weight:800;color:{C['dark']};line-height:1">{avg_phy:.1f}%</p>
      <div style="background:#e5e7eb;border-radius:99px;height:6px;margin-top:8px;overflow:hidden">
        <div style="height:100%;width:{min(avg_phy,100):.0f}%;background:{C['secondary']};border-radius:99px"></div>
      </div>
    </div>
    <div class="sch-kpi">
      <p style="font-size:10px;font-weight:700;letter-spacing:.07em;color:#9ca3af;text-transform:uppercase;margin-bottom:8px">On-Time Delivery</p>
      <p style="font-size:24px;font-weight:800;color:{C['dark']};line-height:1">{on_time_pct}%</p>
      <p style="font-size:11px;color:#6b7280;margin-top:6px">of tracked projects</p>
    </div>
    <div class="sch-kpi">
      <p style="font-size:10px;font-weight:700;letter-spacing:.07em;color:#9ca3af;text-transform:uppercase;margin-bottom:8px">Overdue Items</p>
      <p style="font-size:24px;font-weight:800;color:{C['error']};line-height:1">{overdue_items}</p>
      <p style="font-size:11px;color:{C['error']};margin-top:6px">High severity</p>
    </div>
  </div>

  <!-- Schedule Audit Log -->
  <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;overflow:hidden">
    <div style="padding:14px 20px;border-bottom:1px solid {C['outline']};background:{C['surface_low']};
                display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px">
      <span style="font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:{C['dark']}">Schedule Audit Log</span>
      <div style="display:flex;gap:14px;align-items:center">
        <span style="display:flex;align-items:center;gap:5px;font-size:11px;color:{C['on_var']}">
          <span style="width:8px;height:8px;border-radius:50%;background:{C['secondary']};display:inline-block"></span>On Track
        </span>
        <span style="display:flex;align-items:center;gap:5px;font-size:11px;color:{C['on_var']}">
          <span style="width:8px;height:8px;border-radius:50%;background:{C['error']};display:inline-block"></span>Overdue
        </span>
      </div>
    </div>
    <div style="overflow-x:auto">
      <table style="width:100%;border-collapse:collapse">
        <thead>
          <tr style="background:{C['surface_low']};border-bottom:1px solid {C['outline']}">
            <th style="text-align:left;padding:10px 16px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Project No.</th>
            <th style="text-align:left;padding:10px 16px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Project Name</th>
            <th style="text-align:left;padding:10px 16px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">WO Start</th>
            <th style="text-align:left;padding:10px 16px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">WO Period</th>
            <th style="text-align:left;padding:10px 16px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Ext. Granted</th>
            <th style="text-align:left;padding:10px 16px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Project End Dt.</th>
            <th style="text-align:left;padding:10px 16px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Phy. Compln. %</th>
            <th style="text-align:left;padding:10px 16px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">% Time Elapsed / Overrun</th>
          </tr>
        </thead>
        <tbody>{row_html}</tbody>
      </table>
    </div>
    <div style="padding:14px 20px;border-top:1px solid {C['outline']};display:flex;justify-content:space-between;
                align-items:center;flex-wrap:wrap;gap:10px">
      <span style="font-size:12px;color:{C['on_var']}">Showing {len(page_rows)} of {n_rows} project entries</span>
      <div style="display:flex;gap:6px;align-items:center">
        {'<a href="/schedule' + _qs(page-1) + '" class="sch-pg">‹</a>' if page > 1 else '<span class="sch-pg" style="opacity:.4">‹</span>'}
        {page_links}
        {'<a href="/schedule' + _qs(page+1) + '" class="sch-pg">›</a>' if page < n_pages else '<span class="sch-pg" style="opacity:.4">›</span>'}
      </div>
    </div>
  </div>

</div></div>

<script>
(function(){{
  var PAGE = 'schedule';
  var saved = MonarchFilters.init(PAGE);
  var url   = new URL(window.location.href);
  var hasParams = url.searchParams.get('lob') || url.searchParams.get('pm') || url.searchParams.get('status');
  if (saved && !hasParams) {{
    var qs = [];
    if (saved.lob)    qs.push('lob='    + encodeURIComponent(saved.lob));
    if (saved.pm)     qs.push('pm='     + encodeURIComponent(saved.pm));
    if (saved.status) qs.push('status='+ encodeURIComponent(saved.status));
    if (qs.length) {{ window.location.replace('/schedule?' + qs.join('&')); return; }}
  }}
  var form = document.querySelector('form[action="/schedule"]');
  if (form) {{
    form.addEventListener('submit', function(){{
      var fd = new FormData(form);
      MonarchFilters.save(PAGE, {{ lob: fd.get('lob')||'', pm: fd.get('pm')||'', status: fd.get('status')||'' }});
    }});
  }}
  var qInput = document.getElementById('sch-q-input');
  if (qInput) {{
    var qWasFilled = qInput.value.trim() !== '';
    qInput.addEventListener('input', function(){{
      if (qWasFilled && qInput.value.trim() === '') {{
        form.submit();
      }}
      qWasFilled = qInput.value.trim() !== '';
    }});
  }}
}})();
</script>
</body></html>"""
    return html


# ═══════════════════════════════════════════════════════════════
#  PAGE — EAC STATUS REPORT (next to KPI Monitoring)
# ═══════════════════════════════════════════════════════════════

EAC_STATUS_CHIPS = [
    ("",                   "✓ All Statuses"),
    ("baseline_zero",      "Baseline = 0"),
    ("eac_lt_itd",         "⚠ EAC < ITD Cost"),
    ("wo_lt_itd",          "Work Order < ITD Cost"),
]

EAC_STATUS_BADGES = {
    # key -> (text color, bg color, label)
    "eac_lt_itd":         (C["error"],     "#ffdad6",       "EAC Overrun"),
    "wo_lt_itd":          (C["error"],     "#ffdad6",       "WO Deficit"),
    "baseline_zero":      (C["secondary"], C["sec_cont"],   "No Baseline"),
    "baseline_eac_zero":  (C["secondary"], C["sec_cont"],   "No EAC"),
    "on_track":           (C["on_var"],    C["surface_high"], "On Track"),
}

def _eac_row_metrics(p: dict) -> dict:
    """Estimate-at-Completion variance metrics for one project.

      Work Order Val = Revised Order Value (if set) else Order Value
      EAC             = Rev. Latest EAC (if set) else Baseline EAC
      ITD Cost        = Total expenditure to date
      Estimate to Complete (ETC) = ((1 − %Physical Progress) / %Physical Progress) × ITD Cost
      Remaining %age Phy. Compln. = 100 − Physical Completion %

    Status precedence mirrors the KPI Monitor's RF1 red-flag (EAC < ITD Cost)
    plus the report's own Work-Order and Baseline checks:
      1) EAC Overrun   — EAC entered and EAC < ITD Cost
      2) WO Deficit    — Work Order Value < ITD Cost
      3) No Baseline   — Baseline EAC not yet entered (= 0)
      4) No EAC        — Baseline entered but no EAC entered yet
      5) On Track      — none of the above
    """
    wo_val   = _rep_oi_amt(p)
    eac      = _n(p.get("rev_latest_eac") or p.get("baseline_eac"))
    baseline = _n(p.get("baseline_eac"))
    itd      = _n(p.get("itd_cost"))
    phy      = _rep_phy_pct(p)
    phy_frac = phy / 100.0
    etc      = round(((1 - phy_frac) / phy_frac) * itd, 2) if phy_frac > 0 else 0.0
    rem_phy  = round(max(100 - phy, 0), 1)

    is_eac_overrun = bool(eac > 0 and eac < itd)
    is_wo_deficit  = bool(wo_val > 0 and wo_val < itd)
    is_no_baseline = bool(baseline <= 0)
    is_no_eac      = bool(baseline > 0 and eac <= 0)

    if is_eac_overrun:
        status_key = "eac_lt_itd"
    elif is_wo_deficit:
        status_key = "wo_lt_itd"
    elif is_no_baseline:
        status_key = "baseline_zero"
    elif is_no_eac:
        status_key = "baseline_eac_zero"
    else:
        status_key = "on_track"

    return {
        "wo_val": wo_val, "eac": eac, "baseline": baseline, "itd": itd, "etc": etc,
        "phy": phy, "rem_phy": rem_phy,
        "is_eac_overrun": is_eac_overrun, "is_wo_deficit": is_wo_deficit,
        "is_no_baseline": is_no_baseline, "is_no_eac": is_no_eac,
        "status_key": status_key,
    }

def _eac_filtered_projects(pm="", status="", q="") -> list:
    projects = fetch_all_projects()
    if pm:
        projects = [p for p in projects if (p.get("project_manager") or "") == pm]
    if q:
        qlow = q.strip().lower()
        def _match(p):
            hay = " ".join(str(p.get(k, "") or "") for k in
                            ("id", "co_no", "project_name")).lower()
            return qlow in hay
        projects = [p for p in projects if _match(p)]
    rows = [(p, _eac_row_metrics(p)) for p in projects]
    if status == "baseline_zero":
        rows = [(p, m) for p, m in rows if m["is_no_baseline"]]
    elif status == "baseline_eac_zero":
        rows = [(p, m) for p, m in rows if m["is_no_eac"]]
    elif status == "eac_lt_itd":
        rows = [(p, m) for p, m in rows if m["is_eac_overrun"]]
    elif status == "wo_lt_itd":
        rows = [(p, m) for p, m in rows if m["is_wo_deficit"]]
    return rows

def export_eac_status_excel_bytes(pm="", status="", q="") -> bytes:
    """Export the EAC Status Report — EAC vs ITD Cost variance analysis —
    straight from the DB."""
    from openpyxl.styles import Font, PatternFill, Alignment
    # A "pm" login only ever sees their own projects — ignore any other
    # pm= value that may have been passed (e.g. a hand-edited URL).
    role   = (getattr(_req_ctx, "role", "") or "").lower()
    own_pm = (getattr(_req_ctx, "pm_name", "") or "").strip()
    if role == "pm" and own_pm:
        pm = own_pm
    rows = _eac_filtered_projects(pm, status, q)
    wb = Workbook(); ws = wb.active; ws.title = "EAC Status Report"
    hdr_fill = PatternFill("solid", fgColor="000E24")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    alt_fill = PatternFill("solid", fgColor="F2F4F6")
    headers = ["Co.No.", "Project Name", "LOB", "Project Manager",
               "Work Order Val (₹ Lakhs)", "EAC (₹ Lakhs)", "ITD Cost (₹ Lakhs)",
               "Estimate to Complete (₹ Lakhs)", "Physical Completion", "Status"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = hdr_fill; cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    red_font = Font(color="BA1A1A", bold=True)
    for i, (p, m) in enumerate(rows, start=2):
        label = EAC_STATUS_BADGES.get(m["status_key"], EAC_STATUS_BADGES["on_track"])[2]
        ws.append([
            p.get("co_no", ""), p.get("project_name", ""), _rep_lob(p),
            p.get("project_manager", ""),
            m["wo_val"]/1e5, m["eac"]/1e5, m["itd"]/1e5, m["etc"]/1e5, m["phy"], label,
        ])
        if i % 2 == 0:
            for cell in ws[i]: cell.fill = alt_fill
        if m["is_eac_overrun"] or m["is_wo_deficit"]:
            ws.cell(i, 6).font = red_font
            ws.cell(i, 7).font = red_font
    for idx, w in enumerate([12, 34, 18, 20, 20, 16, 16, 22, 22, 16], start=1):
        ws.column_dimensions[ws.cell(1, idx).column_letter].width = w
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf.read()

def eac_status_report(username="Admin", pm="", status="", q="", page: int = 1) -> str:
    # A "pm" login only ever sees their own projects — ignore any other
    # pm= value that may have been passed (e.g. a hand-edited URL), and
    # don't offer the other-PM options in the selector at all.
    role   = (getattr(_req_ctx, "role", "") or "").lower()
    own_pm = (getattr(_req_ctx, "pm_name", "") or "").strip()
    is_scoped_pm = bool(role == "pm" and own_pm)
    if is_scoped_pm:
        pm = own_pm

    base_rows = _eac_filtered_projects(pm=pm, status="", q=q)   # pm/q-only, for the summary cards
    total_monitored   = len(base_rows)
    critical_overruns = sum(1 for _, m in base_rows if m["is_eac_overrun"])
    missing_baselines = sum(1 for _, m in base_rows if m["is_no_baseline"])
    lob_count = len({_rep_lob(p) for p, _ in base_rows if _rep_lob(p)}) or 1

    rows = _eac_filtered_projects(pm=pm, status=status, q=q)     # status applied too, for the table

    PER_PAGE  = 10
    n_rows    = len(rows)
    n_pages   = max(1, math.ceil(n_rows / PER_PAGE))
    page      = max(1, min(page, n_pages))
    start_ix  = (page - 1) * PER_PAGE
    page_rows = rows[start_ix:start_ix + PER_PAGE]

    def _money(v):
        return f"{v/1e5:,.2f}"

    row_html = ""
    for p, m in page_rows:
        co   = p.get("co_no") or "—"
        name = p.get("project_name") or "—"
        text_col, bg_col, label = EAC_STATUS_BADGES.get(m["status_key"], EAC_STATUS_BADGES["on_track"])

        wo_style  = f"color:{C['error']};font-weight:700" if m["is_wo_deficit"]  else "font-weight:600"
        eac_style = f"color:{C['error']};font-weight:700" if m["is_eac_overrun"] else "font-weight:600"
        itd_style = f"color:{C['error']};font-weight:700" if (m["is_eac_overrun"] or m["is_wo_deficit"]) else "font-weight:600"
        row_bg    = "background:#fff7f6" if m["is_eac_overrun"] else ""

        row_html += f"""
        <tr style="border-bottom:1px solid #f0f2f4;{row_bg}">
          <td style="padding:12px 14px;font-family:'Courier New',monospace;font-size:12px;color:{C['on_var']}">{co}</td>
          <td style="padding:12px 14px;font-size:13px;color:{C['dark']};font-weight:600;max-width:220px;
                     overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="{name}">{name}</td>
          <td style="padding:12px 14px;text-align:right;font-size:13px;{wo_style}">{_money(m['wo_val'])}</td>
          <td style="padding:12px 14px;text-align:right;font-size:13px;{eac_style}">{_money(m['eac'])}</td>
          <td style="padding:12px 14px;text-align:right;font-size:13px;{itd_style}">{_money(m['itd'])}</td>
          <td style="padding:12px 14px;text-align:right;font-size:13px;color:{C['on_surface']}">{_money(m['etc'])}</td>
          <td style="padding:12px 14px;text-align:right;font-size:13px;font-weight:700;color:{C['dark']}">{m['phy']:.0f}%</td>
          <td style="padding:12px 14px;text-align:center">
            <span style="display:inline-flex;align-items:center;padding:3px 10px;border-radius:99px;font-size:10px;
                         font-weight:700;text-transform:uppercase;letter-spacing:.04em;color:{text_col};
                         background:{bg_col};border:1px solid {text_col}33">{label}</span>
          </td>
        </tr>"""

    if not row_html:
        row_html = '<tr><td colspan="8" style="padding:28px;text-align:center;color:#aaa">No projects match these filters.</td></tr>'

    def _qs(pg):
        return (f"?pm={pm}&status={status}&q={q}&page={pg}").replace(" ", "%20")

    page_links = ""
    window_start = max(1, page - 2)
    window_end   = min(n_pages, page + 2)
    if window_start > 1:
        page_links += f'<a href="/eac_status{_qs(1)}" class="sch-pg">1</a><span style="color:#9ca3af">…</span>'
    for pnum in range(window_start, window_end + 1):
        cls = "sch-pg sch-pg-active" if pnum == page else "sch-pg"
        page_links += f'<a href="/eac_status{_qs(pnum)}" class="{cls}">{pnum}</a>'
    if window_end < n_pages:
        page_links += f'<span style="color:#9ca3af">…</span><a href="/eac_status{_qs(n_pages)}" class="sch-pg">{n_pages}</a>'

    if is_scoped_pm:
        pm_selector_html = f"""
      <div style="background:{C['surface_low']};color:{C['dark']};font-size:11px;font-weight:700;
                  letter-spacing:.03em;padding:10px 16px;border-radius:7px;border:1px solid {C['outline']};
                  display:inline-flex;align-items:center;gap:6px">👤 {own_pm}</div>"""
    else:
        pm_opts = "".join(f'<option value="{p_}" {"selected" if p_ == pm else ""}>{p_}</option>'
                           for p_ in get_project_managers())
        # Rendered as a plain <select> (no nested <form>) since it now lives inside
        # the Search + Status Filters form and submits that form on change.
        pm_selector_html = f"""
        <select name="pm" onchange="this.form.submit()"
                style="background:{C['dark']};color:#fff;font-size:11px;font-weight:700;letter-spacing:.03em;
                       padding:10px 30px 10px 16px;border-radius:7px;border:1px solid {C['dark']};
                       appearance:none;cursor:pointer">
          <option value="">Select Project Manager</option>
          {pm_opts}
        </select>"""

    chips_html = ""
    for key, label in EAC_STATUS_CHIPS:
        active = (status == key)
        is_warn = key == "eac_lt_itd"
        if active and is_warn:
            style = f"border:1px solid {C['error']};background:#ffdad64d;color:{C['error']}"
        elif active:
            style = f"border:1px solid {C['dark']};background:{C['dark']}1a;color:{C['dark']}"
        else:
            style = f"border:1px solid {C['outline']};background:#fff;color:{C['on_var']}"
        href = f"/eac_status?pm={pm}&status={key}&q={q}&page=1".replace(" ", "%20")
        chips_html += (f'<a href="{href}" style="{style};padding:8px 16px;border-radius:99px;font-size:11px;'
                        f'font-weight:700;text-decoration:none;display:inline-flex;align-items:center;gap:6px;'
                        f'letter-spacing:.02em;white-space:nowrap">{label}</a>')

    css = f"""<style>
.eac-kpi{{background:#fff;border:1px solid {C['outline']};border-radius:10px;padding:16px 18px}}
.sch-pg{{display:inline-flex;align-items:center;justify-content:center;min-width:30px;height:30px;
         border-radius:7px;border:1px solid {C['outline']};background:#fff;color:{C['on_var']};
         font-size:12px;font-weight:700;text-decoration:none;padding:0 8px}}
.sch-pg:hover{{background:{C['surface_high']}}}
.sch-pg-active{{background:{C['dark']};color:#fff;border-color:{C['dark']}}}
</style>"""

    html  = _html_head("EAC Status Report") + css + "<body>"
    html += _sidebar("eac_status") + _topbar("EAC STATUS REPORT", username)
    html += f"""
<div class="main"><div class="canvas" style="max-width:1360px">

  <div style="display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:20px;flex-wrap:wrap;gap:14px">
    <div>
      <h2 style="font-size:26px;font-weight:800;color:{C['dark']}">EAC Status Report</h2>
      <p style="font-size:13px;color:#6b7280;margin-top:4px">Estimate at Completion (EAC) variance analysis across active projects.</p>
    </div>
    <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
      <a href="/export_excel/eac_status?pm={pm}&status={status}&q={q}"
         style="display:inline-flex;align-items:center;gap:6px;padding:9px 18px;
                background:#fff;border:1px solid {C['outline']};border-radius:8px;font-size:12px;
                font-weight:700;text-decoration:none;color:{C['dark']}">
        ⬇ Export Excel
      </a>
    </div>
  </div>

  <!-- Summary Metrics -->
  <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:20px">
    <div class="eac-kpi">
      <p style="font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['on_var']};margin-bottom:8px">Total Monitored Projects</p>
      <p style="font-size:26px;font-weight:800;color:{C['dark']};line-height:1">{total_monitored:,}</p>
      <p style="font-size:11px;color:#6b7280;margin-top:8px">Across {lob_count} Divisions</p>
    </div>
    <div class="eac-kpi">
      <p style="font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['error']};margin-bottom:8px">Critical Overruns</p>
      <p style="font-size:26px;font-weight:800;color:{C['dark']};line-height:1">{critical_overruns:,}</p>
      <p style="font-size:11px;color:{C['error']};margin-top:8px">↗ EAC &lt; ITD Cost</p>
    </div>
    <div class="eac-kpi">
      <p style="font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['secondary']};margin-bottom:8px">Missing Baselines</p>
      <p style="font-size:26px;font-weight:800;color:{C['dark']};line-height:1">{missing_baselines:,}</p>
      <p style="font-size:11px;color:{C['secondary']};margin-top:8px">⚠ Baseline = 0</p>
    </div>
  </div>

  <!-- Search + Status Filters -->
  <div class="eac-kpi" style="margin-bottom:20px">
    <form method="get" action="/eac_status" style="margin:0 0 14px 0;display:flex;align-items:center;gap:10px;flex-wrap:wrap">
      <input type="hidden" name="status" value="{status}">
      <input type="text" name="q" id="eac-q-input" value="{q.replace('&','&amp;').replace('"','&quot;')}"
             placeholder="🔎 Search by project name or Co.No…"
             style="flex:1;min-width:220px;padding:9px 12px;border:1px solid {C['outline']};border-radius:6px;
                    font-size:13px;background:{C['surface_low']};outline:none;color:{C['dark']}">
      {pm_selector_html}
    </form>
    <p style="font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['on_var']};margin-bottom:10px">Status Filters</p>
    <div style="display:flex;flex-wrap:wrap;gap:8px">
      {chips_html}
    </div>
  </div>

  <!-- Data Table -->
  <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;overflow:hidden">
    <div style="overflow-x:auto">
      <table style="width:100%;border-collapse:collapse;min-width:800px">
        <thead>
          <tr style="background:{C['surface_low']};border-bottom:1px solid {C['outline']}">
            <th style="text-align:left;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Co.No.</th>
            <th style="text-align:left;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Project Name</th>
            <th style="text-align:right;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Work Order Val (₹ Lakhs)</th>
            <th style="text-align:right;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">EAC (₹ Lakhs)</th>
            <th style="text-align:right;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">ITD Cost (₹ Lakhs)</th>
            <th style="text-align:right;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Estimate to Complete (₹ Lakhs)</th>
            <th style="text-align:right;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Physical Completion</th>
            <th style="text-align:center;padding:10px 14px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Status</th>
          </tr>
        </thead>
        <tbody>{row_html}</tbody>
      </table>
    </div>
    <div style="padding:14px 20px;border-top:1px solid {C['outline']};display:flex;justify-content:space-between;
                align-items:center;flex-wrap:wrap;gap:10px">
      <span style="font-size:12px;color:{C['on_var']}">Showing {len(page_rows)} of {n_rows} records</span>
      <div style="display:flex;gap:6px;align-items:center">
        {'<a href="/eac_status' + _qs(page-1) + '" class="sch-pg">‹</a>' if page > 1 else '<span class="sch-pg" style="opacity:.4">‹</span>'}
        {page_links}
        {'<a href="/eac_status' + _qs(page+1) + '" class="sch-pg">›</a>' if page < n_pages else '<span class="sch-pg" style="opacity:.4">›</span>'}
      </div>
    </div>
  </div>

</div></div>

<script>
(function(){{
  var PAGE = 'eac_status';
  var saved = MonarchFilters.init(PAGE);
  var url   = new URL(window.location.href);
  var hasParams = url.searchParams.get('pm') || url.searchParams.get('status');
  if (saved && !hasParams) {{
    var qs = [];
    if (saved.pm)     qs.push('pm='    + encodeURIComponent(saved.pm));
    if (saved.status) qs.push('status='+ encodeURIComponent(saved.status));
    if (qs.length) {{ window.location.replace('/eac_status?' + qs.join('&')); return; }}
  }}
  MonarchFilters.save(PAGE, {{ pm: '{pm}', status: '{status}' }});
  var qInput = document.getElementById('eac-q-input');
  if (qInput) {{
    var qWasFilled = qInput.value.trim() !== '';
    qInput.addEventListener('input', function(){{
      if (qWasFilled && qInput.value.trim() === '') {{
        qInput.closest('form').submit();
      }}
      qWasFilled = qInput.value.trim() !== '';
    }});
    qInput.addEventListener('keydown', function(e){{
      if (e.key === 'Enter') {{ e.preventDefault(); qInput.closest('form').submit(); }}
    }});
  }}
}})();
</script>
</body></html>"""
    return html


# ── Excel export (replaces PDF export) ───────────────────────────────────────

def export_reports_excel_bytes(sections: str = "oi", lob: str = "", search: str = "") -> bytes:
    """
    Export the Consolidated Reports table exactly as currently filtered/viewed:
    parent projects only, matching the LOB filter, search box, and whichever
    metric columns (Order Intake, Collection, Billed, Unbilled, % Expenditure Over Order Value,
    Physical %, Duration) are currently on screen.
    """
    all_projects  = fetch_all_projects()
    projects      = _rep_parents_only(all_projects)         # ── only parent projects ──
    all_children  = _rep_child_projects(all_projects)        # every linked sub-project

    if lob:
        lob_lower = lob.strip().lower()
        projects = [p for p in projects if _rep_lob(p).strip().lower() == lob_lower]

    if search:
        s = search.strip().lower()
        projects = [
            p for p in projects
            if s in (p.get("project_name") or "").lower()
            or s in (p.get("client_name") or "").lower()
            or s in (p.get("co_no") or "").lower()
        ]

    sec_keys = [k.strip() for k in sections.split(",") if k.strip() in REPORT_TAB_META]
    sec_keys = [k for k in REPORT_TAB_ORDER if k in sec_keys] or ["oi"]

    wb = _WB()
    ws = wb.active
    ws.title = "Consolidated Report"[:31]

    # Styles
    hdr_fill   = PatternFill("solid", fgColor="1F2937")
    hdr_font   = Font(bold=True, color="FFFFFF", size=10)
    total_fill = PatternFill("solid", fgColor="F3F4F6")
    total_font = Font(bold=True, size=10)
    amt_font   = Font(name="Courier New", size=10)
    thin       = Side(style="thin", color="E5E7EB")
    border     = Border(left=thin, right=thin, top=thin, bottom=thin)
    right      = Alignment(horizontal="right",  vertical="center")
    left       = Alignment(horizontal="left",   vertical="center")

    n_cols = 4 + len(sec_keys)
    last_col = get_column_letter(n_cols)

    # Title row
    ws.merge_cells(f"A1:{last_col}1")
    ws["A1"] = "Consolidated Report" + (f" — {lob}" if lob else "")
    ws["A1"].font      = Font(bold=True, size=14, color="1F2937")
    ws["A1"].alignment = left
    ws.row_dimensions[1].height = 28

    # Header row
    headers = ["Co. No.", "Project Name", "LOB", "Client"] + [REPORT_TAB_META[k][0] for k in sec_keys]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = left if c <= 4 else right
        cell.border    = border
    ws.row_dimensions[2].height = 20

    # Data rows — each parent project, immediately followed by its sub-projects
    # (indented, lightly shaded) so the tree structure is visible in the sheet.
    totals = {k: 0.0 for k in sec_keys}
    counts = {k: 0 for k in sec_keys}
    r_idx  = 3
    child_fill = PatternFill("solid", fgColor="F4F6FB")
    child_font = Font(size=10, italic=True, color="4B5563")

    def _write_metric_cells(row_idx, p, fill, font, count_totals):
        for i, k in enumerate(sec_keys):
            col  = 5 + i
            typ  = REPORT_TAB_META[k][2]
            val  = REPORT_TAB_META[k][3](p)
            cell = ws.cell(row=row_idx, column=col)
            cell.fill, cell.border, cell.alignment = fill, border, right
            cell.font = font if font else amt_font
            if val is None:
                cell.value = "—"
            elif typ == "currency":
                cell.value = val
                cell.number_format = '₹#,##0.00'
                if count_totals: totals[k] += val; counts[k] += 1
            elif typ == "percent":
                cell.value = val / 100.0
                cell.number_format = '0.0%'
                if count_totals: totals[k] += val; counts[k] += 1
            elif typ == "days":
                cell.value = val
                cell.number_format = '0"d"'
                if count_totals: totals[k] += val; counts[k] += 1

    for p in projects:
        fill = PatternFill("solid", fgColor="FAFAFA") if r_idx % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        row_vals = [p.get("co_no") or "—", p.get("project_name") or "—", _rep_lob(p) or "—", p.get("client_name") or "—"]
        for c, val in enumerate(row_vals, 1):
            cell = ws.cell(row=r_idx, column=c, value=val)
            cell.fill, cell.border, cell.alignment = fill, border, left
            cell.font = Font(size=10)
        # totals/averages are computed from parent projects only, to avoid
        # double-counting figures that already roll up from sub-projects
        _write_metric_cells(r_idx, p, fill, None, count_totals=True)
        r_idx += 1

        kids = _rep_kids_of(p, all_children)
        for c_p in kids:
            ws.cell(row=r_idx, column=1, value=c_p.get("co_no") or "—")
            ws.cell(row=r_idx, column=2, value="   ↳ " + (c_p.get("project_name") or "—"))
            ws.cell(row=r_idx, column=3, value=_rep_lob(c_p) or "—")
            ws.cell(row=r_idx, column=4, value=c_p.get("client_name") or "—")
            for col in (1, 2, 3, 4):
                cell = ws.cell(row=r_idx, column=col)
                cell.fill, cell.border, cell.alignment, cell.font = child_fill, border, left, child_font
            _write_metric_cells(r_idx, c_p, child_fill, child_font, count_totals=False)
            r_idx += 1

    # Total / Average footer row
    footer_row = r_idx
    ws.merge_cells(f"A{footer_row}:D{footer_row}")
    ws[f"A{footer_row}"]           = "TOTAL / AVERAGE"
    ws[f"A{footer_row}"].font      = total_font
    ws[f"A{footer_row}"].fill      = total_fill
    ws[f"A{footer_row}"].alignment = left
    ws[f"A{footer_row}"].border    = border
    for i, k in enumerate(sec_keys):
        col  = 5 + i
        typ  = REPORT_TAB_META[k][2]
        cell = ws.cell(row=footer_row, column=col)
        cell.font, cell.fill, cell.alignment, cell.border = (
            Font(bold=True, name="Courier New", size=11), total_fill, right, border
        )
        if typ == "currency":
            cell.value = totals[k]
            cell.number_format = '₹#,##0.00'
        elif typ == "percent":
            cell.value = (totals[k] / counts[k] / 100.0) if counts[k] else 0
            cell.number_format = '0.0%'
        elif typ == "days":
            cell.value = round(totals[k] / counts[k]) if counts[k] else 0
            cell.number_format = '0"d"'

    # Column widths
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22
    for i in range(len(sec_keys)):
        ws.column_dimensions[get_column_letter(5 + i)].width = 20

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ── Main reports page ─────────────────────────────────────────────────────────
def reports(alert_html="", username="Admin", active_section="") -> str:
    """
    Consolidated Reports page — single sidebar entry, no sub-items.
    Filter bar: LOB picker + 4 report-type tabs.
    Export button downloads Excel.
    """
    lob_list      = get_visible_lob_list()
    all_projects  = fetch_all_projects()
    projects      = _rep_parents_only(all_projects)      # ── parent projects shown as top-level rows ──
    all_children  = _rep_child_projects(all_projects)     # every linked sub-project
 
    now        = datetime.now().strftime("%d %b %Y, %I:%M %p IST")
    active_tab = active_section if active_section in ("oi", "col", "itd", "bnc", "ub", "fin", "phy", "dur") else "oi"
 
    # Build all project rows as JSON for client-side filtering.
    # Each parent row carries its "children" (sub-projects) so the "Sub Projects"
    # view can list them alongside their parent.
    rows_data = []
    for p in projects:
        row = _rep_row(p)
        kids = _rep_kids_of(p, all_children)
        row["children"] = [_rep_row(c) for c in kids]
        rows_data.append(row)
 
    rows_json     = _json.dumps(rows_data, ensure_ascii=False)
    lob_opts_json = _json.dumps(lob_list or [], ensure_ascii=False)
 
    html  = _html_head("Consolidated Reports") + "<body>"
    # Pass no active_section sub-key — sidebar shows only "Consolidated Reports"
    html += _sidebar("reports") + _topbar("CONSOLIDATED REPORTS", username)
    html += f"""
<style>
.rep-bar {{
    display:flex; align-items:center; gap:12px; flex-wrap:wrap;
    background:#fff; border:1px solid {C['outline']}; border-radius:12px;
    padding:14px 18px; margin-bottom:18px;
    box-shadow:0 1px 2px rgba(16,24,40,.04);
}}
.rep-lob-label {{
    font-size:11px; font-weight:700; letter-spacing:.06em;
    text-transform:uppercase; color:{C['on_var']}; white-space:nowrap;
}}
.rep-lob-select, .rep-search {{
    padding:8px 12px; border:1px solid {C['outline']}; border-radius:7px;
    font-size:12px; background:{C['surface_low']}; color:{C['dark']};
    outline:none; cursor:pointer;
}}
.rep-lob-select {{ min-width:150px; }}
.rep-search {{ min-width:200px; cursor:text; }}
.rep-search:focus, .rep-lob-select:focus {{ border-color:{C['dark']}; background:#fff; }}
.rep-divider {{ width:1px; height:26px; background:{C['outline']}; flex-shrink:0; }}
.rep-tabs    {{ display:flex; gap:6px; flex-wrap:wrap; align-items:center; }}
.rep-tab {{
    padding:7px 16px; border-radius:20px; border:1px solid {C['outline']};
    font-size:12px; font-weight:600; cursor:pointer;
    background:{C['surface_low']}; color:{C['on_var']};
    transition:all .15s; white-space:nowrap; display:inline-flex; align-items:center; gap:6px;
}}
.rep-tab::before {{
    content:''; width:7px; height:7px; border-radius:50%;
    background:{C['outline']}; transition:background .15s;
}}
.rep-tab:hover  {{ background:{C['surface_high']}; color:{C['dark']}; }}
.rep-tab.active {{ background:{C['dark']}; color:#fff; border-color:{C['dark']}; }}
.rep-tab.active::before {{ background:#22c55e; }}
.rep-mini-link {{
    font-size:11px; font-weight:700; color:{C['secondary']}; cursor:pointer;
    text-decoration:underline; white-space:nowrap; background:none; border:none;
}}
.rep-card {{
    background:#fff; border:1px solid {C['outline']};
    border-radius:14px; overflow:hidden; margin-bottom:20px;
    box-shadow:0 2px 8px rgba(16,24,40,.06);
}}
.rep-card-head {{
    background:linear-gradient(135deg, {C['dark']} 0%, #1f2937 100%);
    padding:14px 20px;
    display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px;
}}
.rep-tbl-wrap {{ max-height:600px; overflow:auto; }}
.rep-tbl {{ width:100%; border-collapse:collapse; font-size:13px; white-space:nowrap; }}
.rep-tbl thead th {{
    position:sticky; top:0; z-index:2;
    padding:11px 16px; font-size:10px; font-weight:800;
    letter-spacing:.06em; text-transform:uppercase; color:{C['on_var']}; text-align:left;
    background:{C['surface_low']};
    box-shadow:inset 0 -1px 0 {C['outline']};
}}
.rep-tbl thead th.fixed-col {{ z-index:3; background:{C['surface_low']}; }}
.rep-tbl tbody td.fixed-col {{ background:inherit; }}

/* ── Vertical freeze pane: Co.No + Project Name stay put while the
   rest of the table scrolls horizontally underneath them ── */
.rep-tbl-wrap {{ overflow-x:auto; }}
.rep-freeze-1 {{ position:sticky; left:0;    z-index:5; }}
.rep-freeze-2 {{ position:sticky; left:110px; z-index:5; box-shadow:2px 0 4px -2px rgba(0,0,0,.15); }}
thead .rep-freeze-1, thead .rep-freeze-2 {{ z-index:6; background:{C['surface_low']}; }}
tbody .rep-freeze-1, tbody .rep-freeze-2 {{ background:#fff; }}
tbody tr:nth-child(even) .rep-freeze-1,
tbody tr:nth-child(even) .rep-freeze-2 {{ background:#fafbfc; }}
tbody tr:hover .rep-freeze-1,
tbody tr:hover .rep-freeze-2 {{ background:#eef2ff; }}
tfoot .rep-freeze-1, tfoot .rep-freeze-2 {{ z-index:6; background:{C['surface_high']}; }}
.rep-tbl tbody tr {{ transition:background .12s; }}
.rep-tbl tbody tr:nth-child(even) {{ background:#fafbfc; }}
.rep-tbl tbody tr:hover {{ background:#eef2ff !important; }}
.rep-tbl td {{ padding:11px 16px; border-bottom:1px solid #f1f2f4; color:{C['dark']}; }}
.rep-tbl tfoot td {{
    position:sticky; bottom:0; z-index:2;
    padding:12px 16px; font-family:'Courier New',monospace; font-weight:800; font-size:13px;
    color:{C['dark']}; background:{C['surface_high']};
    box-shadow:inset 0 1px 0 {C['outline']};
    text-align:right;
}}
.rep-tbl tfoot td.foot-label {{
    text-align:left; font-family:inherit; font-size:11px; font-weight:800;
    letter-spacing:.05em; text-transform:uppercase; color:{C['dark']};
}}
.pct-chip {{
    display:inline-block; padding:3px 10px; border-radius:20px;
    font-weight:700; font-size:12px; font-family:'Courier New',monospace;
}}
.pct-good {{ background:#dcfce7; color:#166534; }}
.pct-mid  {{ background:#fef3c7; color:#92400e; }}
.pct-bad  {{ background:#fee2e2; color:#991b1b; }}
#no-rows {{
    display:none; padding:32px; text-align:center; color:#9ca3af; font-size:13px;
}}
.export-xl {{
    display:inline-flex; align-items:center; gap:6px;
    padding:8px 16px; border-radius:7px; font-size:12px; font-weight:600;
    background:#166534; color:#fff; text-decoration:none; border:none; cursor:pointer;
    transition:background .15s;
}}
.export-xl:hover {{ background:#14532d; }}
.rep-count-badge {{
    display:inline-flex; align-items:center; padding:4px 10px; border-radius:20px;
    background:rgba(255,255,255,.12); color:#fff; font-size:11px; font-weight:700;
}}
</style>
 
<div class="main"><div class="canvas" style="max-width:1300px">
  {alert_html}
 
  <!-- Page header -->
  <div style="display:flex;justify-content:space-between;align-items:flex-end;
              margin-bottom:20px;flex-wrap:wrap;gap:12px">
    <div>
      <p style="font-size:11px;font-weight:600;letter-spacing:.07em;text-transform:uppercase;
                color:{C['secondary']};margin:0 0 4px">Corporate Performance Suite</p>
      <h2 style="font-size:26px;font-weight:700;color:{C['dark']};margin:0"
          id="page-title">Consolidated Reports</h2>
      <p style="font-size:11px;color:{C['on_var']};margin:6px 0 0">
        Click <span style="color:{C['secondary']};font-weight:700">Sub Projects</span> in the bar below
        to view every sub-project along with its parent project
      </p>
    </div>
    <a href="/export_excel/reports?section={active_tab}" class="export-xl" id="export-btn">
      ⬇ Export Excel
    </a>
  </div>
 
  <!-- Filter bar: search + LOB + toggle tabs -->
  <div class="rep-bar">
    <input class="rep-search" id="proj-search" type="text"
           placeholder="🔎 Search project, co. no. or client…" oninput="renderRows()">
    <div class="rep-divider"></div>
    <span class="rep-lob-label">LOB</span>
    <select class="rep-lob-select" id="lob-filter" onchange="renderRows()">
      <option value="">All LOB</option>
    </select>
    <div class="rep-divider"></div>
    <button class="rep-tab" id="subproj-toggle" onclick="toggleSubView()">🌿 Sub Projects</button>
    <div class="rep-divider"></div>
    <div class="rep-tabs">
      <button class="rep-tab" data-tab="oi"  onclick="toggleTab('oi')">Order Intake</button>
      <button class="rep-tab" data-tab="col" onclick="toggleTab('col')">Collection</button>
      <button class="rep-tab" data-tab="itd" onclick="toggleTab('itd')">Billed Amount</button>
      <button class="rep-tab" data-tab="bnc" onclick="toggleTab('bnc')">Billed but Not Collected</button>
      <button class="rep-tab" data-tab="ub"  onclick="toggleTab('ub')">Unbilled Amount</button>
      <button class="rep-tab" data-tab="fin" onclick="toggleTab('fin')">% Expenditure Over Order Value</button>
      <button class="rep-tab" data-tab="phy" onclick="toggleTab('phy')">Physical %</button>
      <button class="rep-tab" data-tab="dur" onclick="toggleTab('dur')">Duration</button>
    </div>
    <button class="rep-mini-link" onclick="resetColumns()" style="margin-left:auto">Reset columns</button>
  </div>
 
  <!-- Report card -->
  <div class="rep-card">
    <div class="rep-card-head">
      <span style="font-size:13px;font-weight:700;letter-spacing:.03em;
                   text-transform:uppercase;color:#fff" id="card-title">Consolidated Report</span>
      <span class="rep-count-badge" id="card-tag">0 columns · 0 projects</span>
    </div>
    <div class="rep-tbl-wrap">
      <table class="rep-tbl">
        <thead>
          <tr id="head-row">
        <th class="fixed-col rep-freeze-1" style="min-width:110px">Co. No.</th>
        <th class="fixed-col rep-freeze-2" style="min-width:260px">Project Name</th>
        <th class="fixed-col" style="min-width:220px;display:none" id="th-parent">Parent Project</th>
        <th class="fixed-col" style="min-width:130px">LOB</th>
        <th class="fixed-col" style="min-width:150px">Client</th>
</tr>
        </thead>
        <tbody id="rep-tbody"></tbody>
        <tfoot>
          <tr id="foot-row">
            <td class="foot-label" colspan="4" id="foot-label-cell">TOTAL / AVERAGE</td>
          </tr>
        </tfoot>
      </table>
      <div id="no-rows">No matching projects.</div>
    </div>
  </div>
 
  <div style="margin-top:12px;padding:10px 0;border-top:1px solid {C['outline']};
              display:flex;justify-content:space-between;font-size:11px;color:{C['on_var']}">
    <span>© 2025 Monarch Surveyors &amp; Engineering Consultants Ltd.</span>
    <span>🟢 Live DB &nbsp;|&nbsp; {now}</span>
  </div>
</div></div>
 
<script>
const ALL_ROWS = {rows_json};
const LOB_LIST = {lob_opts_json};
 
// Fixed left-to-right order for dynamic columns, regardless of click order
const TAB_ORDER = ["oi", "col", "itd", "bnc", "ub", "fin", "phy", "dur"];
const TAB_META = {{
  oi:  {{ label:"Order Intake",    header:"Order Intake (₹)",    footer:"Total Order Intake",    key:"oi",  type:"currency" }},
  col: {{ label:"Collection",      header:"Collection (₹)",      footer:"Total Collection",       key:"col", type:"currency" }},
  itd: {{ label:"Billed Amount",   header:"Billed Amount (₹)",   footer:"Total Billed Amount",    key:"itd", type:"currency" }},
  bnc: {{ label:"Billed but Not Collected", header:"Billed but Not Collected (₹)", footer:"Total Billed but Not Collected", key:"bnc", type:"currency" }},
  ub:  {{ label:"Unbilled Amount", header:"Unbilled Amount (₹)", footer:"Total Unbilled Amount",  key:"ub",  type:"currency" }},
  fin: {{ label:"% Expenditure Over Order Value", header:"% Expenditure Over Order Value", footer:"Avg. % Expenditure Over Order Value", key:"fin", type:"percent"  }},
  phy: {{ label:"Physical %",      header:"Physical %",          footer:"Avg. Physical %",        key:"phy", type:"percent"  }},
  dur: {{ label:"Duration",        header:"Duration",            footer:"Avg. Duration",          key:"dur", type:"days"     }},
}};
const DEFAULT_TABS = TAB_ORDER.slice(); // show every metric column by default
 
// Tabs that are currently "added" as columns. All shown by default so nothing is hidden.
let activeTabs = new Set(DEFAULT_TABS);
 
// "Sub Projects" view: flips the table from a plain parent-projects list
// into a flat, read-only list of every sub-project alongside its parent's details.
let showSubProjects = false;
function toggleSubView() {{
  showSubProjects = !showSubProjects;
  document.getElementById('subproj-toggle').classList.toggle('active', showSubProjects);
  document.getElementById('card-title').textContent = showSubProjects ? 'Sub Projects' : 'Consolidated Report';
  document.getElementById('th-parent').style.display = showSubProjects ? '' : 'none';
  document.getElementById('foot-label-cell').colSpan = showSubProjects ? 5 : 4;
  renderRows();
}}
 
// Populate LOB dropdown from DB values
const lobSel = document.getElementById('lob-filter');
LOB_LIST.forEach(function(lb) {{
  if (!lb) return;
  const o = document.createElement('option');
  o.value = lb; o.textContent = lb;
  lobSel.appendChild(o);
}});
 
function fmtCurrency(v) {{
  if (Math.abs(v) >= 1e7) return '₹\u200a' + (v/1e7).toFixed(2) + '\u00a0Cr';
  if (Math.abs(v) >= 1e5) return '₹\u200a' + (v/1e5).toFixed(2) + '\u00a0L';
  return '₹\u200a' + v.toLocaleString('en-IN', {{maximumFractionDigits:0}});
}}
function fmtDays(v) {{
  if (v === null || v === undefined) return '—';
  return v < 0 ? ('Overdue ' + Math.abs(v) + 'd') : (v + 'd left');
}}
// Percent / duration values render as colour-coded chips so status is readable at a glance
function pctChip(v) {{
  if (v === null || v === undefined) return '—';
  const cls = v >= 75 ? 'pct-good' : (v >= 40 ? 'pct-mid' : 'pct-bad');
  return '<span class="pct-chip ' + cls + '">' + v.toFixed(1) + '%</span>';
}}
function durChip(v) {{
  if (v === null || v === undefined) return '—';
  const cls = v < 0 ? 'pct-bad' : (v <= 30 ? 'pct-mid' : 'pct-good');
  return '<span class="pct-chip ' + cls + '">' + fmtDays(v) + '</span>';
}}
function cellByType(type, v) {{
  if (type === 'percent') return pctChip(v);
  if (type === 'days')    return durChip(v);
  return fmtCurrency(v || 0);
}}
 
function orderedActiveTabs() {{
  return TAB_ORDER.filter(function(t) {{ return activeTabs.has(t); }});
}}
 
// Clicking a tab ADDS its column immediately; clicking an already-active tab removes it
// (at least one column always stays visible).
function toggleTab(tab) {{
  if (activeTabs.has(tab)) {{
    if (activeTabs.size > 1) activeTabs.delete(tab);
  }} else {{
    activeTabs.add(tab);
  }}
  syncTabButtons();
  renderHeader();
  renderRows();
}}
 
function resetColumns() {{
  activeTabs = new Set(DEFAULT_TABS);
  syncTabButtons();
  renderHeader();
  renderRows();
}}
 
function syncTabButtons() {{
  document.querySelectorAll('.rep-tab').forEach(function(b) {{
    b.classList.toggle('active', activeTabs.has(b.dataset.tab));
  }});
  updateExportLink();
}}
 
// Keeps the Export Excel button in sync with whatever is currently on screen:
// selected columns + LOB filter + search text.
function updateExportLink() {{
  const tabs   = orderedActiveTabs();
  const lob    = lobSel.value;
  const search = (document.getElementById('proj-search').value || '').trim();
  let url = '/export_excel/reports?section=' + encodeURIComponent(tabs.join(','));
  if (lob)    url += '&lob=' + encodeURIComponent(lob);
  if (search) url += '&q='   + encodeURIComponent(search);
  document.getElementById('export-btn').href = url;
}}
 
function renderHeader() {{
  const headRow = document.getElementById('head-row');
  const footRow = document.getElementById('foot-row');
  // Remove previously injected dynamic columns, keep the 3 fixed ones
  headRow.querySelectorAll('th[data-dyn]').forEach(function(th) {{ th.remove(); }});
  footRow.querySelectorAll('td[data-dyn]').forEach(function(td) {{ td.remove(); }});
  orderedActiveTabs().forEach(function(t) {{
    const th = document.createElement('th');
    th.setAttribute('data-dyn', t);
    th.style.textAlign  = 'right';
    th.style.minWidth   = '150px';
    th.textContent = TAB_META[t].header;
    headRow.appendChild(th);
 
    const td = document.createElement('td');
    td.setAttribute('data-dyn', t);
    td.id = 'foot-' + t;
    footRow.appendChild(td);
  }});
}}
 
function renderRows() {{
  const lob    = lobSel.value;
  const search = (document.getElementById('proj-search').value || '').trim().toLowerCase();
  const tabs   = orderedActiveTabs();
  updateExportLink();
  MonarchFilters.save('reports', {{
    tabs: orderedActiveTabs(), lob: lobSel.value,
    search: document.getElementById('proj-search').value, showSub: showSubProjects
  }});
  const tbody  = document.getElementById('rep-tbody');
  tbody.innerHTML = '';
  const totals = {{}}; const counts = {{}};
  tabs.forEach(function(t) {{ totals[t] = 0; counts[t] = 0; }});
  let count = 0;
 
  function textMatch(row, needle) {{
    return (row.name||'').toLowerCase().includes(needle)
        || (row.client||'').toLowerCase().includes(needle)
        || (row.co_no||'').toLowerCase().includes(needle);
  }}

  if (showSubProjects) {{
    // Flat, view-only list: every sub-project, with its parent's Co. No. / name shown alongside it.
    ALL_ROWS.forEach(function(r) {{
      if (lob && r.lob !== lob) return;
      (r.children || []).forEach(function(k) {{
        const kidMatches    = !search || textMatch(k, search);
        const parentMatches = !search || textMatch(r, search);
        if (!kidMatches && !parentMatches) return;
        count++;
        const tr = document.createElement('tr');
        let cells =
          '<td class="fixed-col" style="font-size:12px;color:#6b7280">'  + (k.co_no || '—') + '</td>' +
          '<td class="fixed-col" style="font-weight:600;color:#111827">' + k.name            + '</td>' +
          '<td class="fixed-col" style="font-size:12px;color:#6b7280">'
            + (r.co_no && r.co_no !== '—' ? r.co_no + ' — ' : '') + r.name + '</td>' +
          '<td class="fixed-col" style="font-size:12px;color:#6b7280">'  + (k.lob  || '—')  + '</td>' +
          '<td class="fixed-col" style="color:#6b7280">'                 + (k.client || '—') + '</td>';
        tabs.forEach(function(t) {{
          const type = TAB_META[t].type;
          const val  = k[TAB_META[t].key];
          if (val !== null && val !== undefined) {{ totals[t] += val; counts[t] += 1; }}
          cells += '<td style="text-align:right">' + cellByType(type, val) + '</td>';
        }});
        tr.innerHTML = cells;
        tbody.appendChild(tr);
      }});
    }});
    document.getElementById('no-rows').style.display = count ? 'none' : 'block';
    document.getElementById('card-tag').textContent =
      tabs.length + ' column' + (tabs.length===1?'':'s') + ' · ' + count + ' sub-project' + (count===1?'':'s');
  }} else {{
    // Default view: plain list of parent projects (no tree/expand — use the
    // "Sub Projects" button above to see sub-projects with their parent).
    ALL_ROWS.forEach(function(r) {{
      if (lob && r.lob !== lob) return;
      if (search && !textMatch(r, search)) return;
      count++;

      const tr = document.createElement('tr');
      let cells =
        '<td class="fixed-col" style="font-size:12px;color:#6b7280">'  + (r.co_no || '—') + '</td>' +
        '<td class="fixed-col" style="font-weight:600;color:#111827">' + r.name + '</td>' +
        '<td class="fixed-col" style="font-size:12px;color:#6b7280">'  + (r.lob  || '—')  + '</td>' +
        '<td class="fixed-col" style="color:#6b7280">'                 + (r.client || '—') + '</td>';
      tabs.forEach(function(t) {{
        const type = TAB_META[t].type;
        const val  = r[TAB_META[t].key];
        // Totals reflect parent projects only, so sub-project figures already
        // rolled into a parent aren't double-counted.
        if (val !== null && val !== undefined) {{ totals[t] += val; counts[t] += 1; }}
        cells += '<td style="text-align:right">' + cellByType(type, val) + '</td>';
      }});
      tr.innerHTML = cells;
      tbody.appendChild(tr);
    }});
    document.getElementById('no-rows').style.display = count ? 'none' : 'block';
    document.getElementById('card-tag').textContent =
      tabs.length + ' column' + (tabs.length===1?'':'s') + ' · ' + count + ' project' + (count===1?'':'s');
  }}
 
  // Totals row: lines up exactly under its column because it shares the same <table>
  tabs.forEach(function(t) {{
    const type  = TAB_META[t].type;
    const isAvg = (type === 'percent' || type === 'days');
    const value = isAvg ? (counts[t] ? totals[t] / counts[t] : null) : totals[t];
    const cell  = document.getElementById('foot-' + t);
    if (!cell) return;
    cell.textContent = isAvg
      ? (type === 'percent' ? (value===null?'—':value.toFixed(1)+'%') : fmtDays(value===null?null:Math.round(value)))
      : fmtCurrency(value);
  }});
}}
 
// Initialise on load
(function(){{
  var saved = MonarchFilters.init('reports');
  if (saved) {{
    if (Array.isArray(saved.tabs) && saved.tabs.length) {{
      activeTabs = new Set(saved.tabs.filter(function(t){{ return TAB_ORDER.indexOf(t) !== -1; }}));
    }}
    if (saved.lob)    lobSel.value = saved.lob;
    if (saved.search) document.getElementById('proj-search').value = saved.search;
    if (saved.showSub) {{
      showSubProjects = true;
      document.getElementById('subproj-toggle').classList.add('active');
      document.getElementById('card-title').textContent = 'Sub Projects';
      document.getElementById('th-parent').style.display = '';
      document.getElementById('foot-label-cell').colSpan = 5;
    }}
  }}
}})();
syncTabButtons();
renderHeader();
renderRows();
</script>
</body></html>"""
    return html
 
def fetch_projects_for_select():
    try:
        with ENGINE.connect() as conn:
            rows = conn.execute(text(
                "SELECT id, co_no, project_name, client_name, order_number, order_value, "
                "business_line, variation_amt, revised_order_value, hier_business_line, "
                "project_manager "
                "FROM projects ORDER BY id DESC"
            )).fetchall()
        return rows
    except Exception:
        return []
 
def fetch_projects_for_lob(lob: str):
    if not lob:
        return []
    all_rows = fetch_projects_for_select()
    lob_lower = lob.strip().lower()
    return [
        r for r in all_rows
        if str(r[6] or "").strip().lower() == lob_lower
        or str(r[9] or "").strip().lower() == lob_lower
    ]

def fetch_projects_for_pm(pm_name: str):
    """Same row shape as fetch_projects_for_select(), filtered to one PM's projects."""
    if not pm_name:
        return []
    try:
        with ENGINE.connect() as conn:
            rows = conn.execute(text(
                "SELECT id, co_no, project_name, client_name, order_number, order_value, "
                "business_line, variation_amt, revised_order_value, hier_business_line "
                "FROM projects WHERE project_manager=:pm ORDER BY id DESC"
            ), {"pm": pm_name}).fetchall()
        return rows
    except Exception as e:
        print(f"[fetch_projects_for_pm] {e}")
        return []

def pm_owns_project(pm_name: str, project_id) -> bool:
    """True if `project_id` belongs to a project managed by `pm_name`. Used to
    stop a scoped 'pm' login from mutating another PM's billing entries by
    hand-crafting a POST with someone else's project_id/entry_id."""
    pm_name = (pm_name or "").strip()
    if not pm_name or not project_id:
        return False
    try:
        with ENGINE.connect() as conn:
            row = conn.execute(text(
                "SELECT 1 FROM projects WHERE id=:id AND LOWER(TRIM(project_manager))=:pm"
            ), {"id": project_id, "pm": pm_name.lower()}).fetchone()
        return bool(row)
    except Exception as e:
        print(f"[pm_owns_project] {e}")
        return False

def pm_owns_billing_entry(pm_name: str, entry_id) -> bool:
    """True if the billing_collection row `entry_id` belongs to a project
    managed by `pm_name`."""
    pm_name = (pm_name or "").strip()
    if not pm_name or not entry_id:
        return False
    try:
        with ENGINE.connect() as conn:
            row = conn.execute(text("""
                SELECT 1 FROM billing_collection bc
                JOIN projects p ON p.id = bc.project_id
                WHERE bc.id=:eid AND LOWER(TRIM(p.project_manager))=:pm
            """), {"eid": entry_id, "pm": pm_name.lower()}).fetchone()
        return bool(row)
    except Exception as e:
        print(f"[pm_owns_billing_entry] {e}")
        return False

# ═══════════════════════════════════════════════════════════════
#  PAGE 8— BILLING & COLLECTIONS
# ═══════════════════════════════════════════════════════════════
def user_management(alert_html="", username="Admin") -> str:
    users = fetch_all_users()

    rows_html = ""
    for u in users:
        is_protected = u['username'] == 'admin'
        cur_role = (u['role'] or 'user').lower()
        pw_display = (u.get('password_hash') or '')[:12] + "…" if u.get('password_hash') else "—"

        if is_protected:
            role_cell = (
                f'<span style="padding:3px 10px;border-radius:12px;font-size:11px;font-weight:700;'
                f'background:#fef3c7;color:#92400e">SUPERADMIN</span>'
            )
        else:
            role_opts = ""
            for rv in ["user", "director", "pm", "admin"]:
                sel = 'selected' if cur_role == rv else ''
                role_opts += f'<option value="{rv}" {sel}>{rv.upper()}</option>'
            role_cell = (
                f'<form method="post" action="/update_user_role" style="display:inline;margin:0">'
                f'<input type="hidden" name="user_id" value="{u["id"]}">'
                f'<select name="role" onchange="this.form.submit()" '
                f'style="padding:4px 8px;border:1px solid {C["outline"]};border-radius:6px;'
                f'font-size:12px;font-weight:600;background:#fff;cursor:pointer;outline:none">'
                f'{role_opts}</select></form>'
            )

        actions_html = ""
        if not is_protected:
            actions_html = (
                f'<button type="button" class="btn-outline" style="padding:4px 10px;font-size:11px;margin-right:4px" '
                f'onclick="document.getElementById(\'rst-{u["id"]}\').style.display=\'table-row\'" '
                f'title="Reset password">🔑 Reset</button>'
                f'<form method="post" action="/delete_user" style="display:inline;margin:0" '
                f'onsubmit="return confirm(\'Delete user &quot;{u["username"]}&quot;? This cannot be undone.\')">'
                f'<input type="hidden" name="user_id" value="{u["id"]}">'
                f'<button class="del-btn" title="Delete user">🗑</button>'
                f'</form>'
            )

        reset_row = ""
        if not is_protected:
            reset_row = (
                f'<tr id="rst-{u["id"]}" style="display:none;background:#fffbeb">'
                f'<td colspan="6" style="padding:10px 16px">'
                f'<form method="post" action="/reset_user_password" '
                f'style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">'
                f'<input type="hidden" name="username" value="{u["username"]}">'
                f'<span style="font-size:12px;font-weight:600;color:{C["dark"]}">'
                f'New password for <b>{u["username"]}</b>:</span>'
                f'<input type="password" name="new_password" placeholder="Min. 6 chars" required minlength="6" '
                f'style="padding:5px 10px;border:1px solid {C["outline"]};border-radius:6px;'
                f'font-size:12px;width:180px">'
                f'<button type="submit" class="btn-primary" style="padding:5px 14px;font-size:11px">'
                f'Save Password</button>'
                f'<button type="button" class="btn-outline" style="padding:5px 12px;font-size:11px" '
                f'onclick="this.closest(\'tr\').style.display=\'none\'">Cancel</button>'
                f'</form></td></tr>'
            )

        created_display = (u['created_at'] or '')[:19].replace('T', ' ') or '—'

        rows_html += f"""
        <tr style="border-bottom:1px solid {C['outline']}">
          <td style="padding:12px 16px;font-weight:700;color:{C['on_var']};font-size:12px">{u['id']}</td>
          <td style="padding:12px 16px;font-weight:600;color:{C['dark']}">{u['username']}</td>
          <td style="padding:12px 16px;font-family:monospace;font-size:11px;color:{C['on_var']}">{pw_display}</td>
          <td style="padding:12px 16px;color:{C['on_var']}">{u['display_name'] or '—'}</td>
          <td style="padding:12px 16px">{role_cell}</td>
          <td style="padding:12px 16px;color:{C['on_var']};font-size:12px">{u.get('pm_name') or '—'}</td>
          <td style="padding:12px 16px;color:{C['on_var']};font-size:12px">{created_display}</td>
          <td style="padding:12px 16px;text-align:right;white-space:nowrap">{actions_html}</td>
        </tr>
        {reset_row}"""

        if not rows_html:
          rows_html = '<tr><td colspan="8" style="padding:20px;text-align:center;color:#aaa">No users yet.</td></tr>'

    html  = _html_head("User Management") + "<body>"
    html += _sidebar("user_management", is_admin=True) + _topbar("USER MANAGEMENT", username)
    html += f"""
<div class="main"><div class="canvas" style="max-width:1100px">
  <h2 style="font-size:24px;font-weight:800;color:{C['dark']};margin-bottom:4px">User Management</h2>
  <p style="font-size:13px;color:#6b7280;margin-bottom:20px">Add, edit, or remove users who can log in to this system.</p>
  {alert_html}

   <div class="card" style="margin-bottom:24px">
    <div style="padding:20px 24px;display:flex;justify-content:space-between;align-items:center;gap:16px;flex-wrap:wrap">
      <div>
        <div class="sec-title" style="margin-bottom:4px">🧑‍💼 Independent PM Logins</div>
        <p style="font-size:12px;color:#6b7280">
          Auto-create a separate login for every Project Manager who doesn't have one yet.
          Each PM will only see their own projects across the whole system.
        </p>
      </div>
      <form method="post" action="/generate_pm_logins">
        <button type="submit" class="btn-primary">⚙ Generate Logins for All PMs</button>
      </form>
    </div>
  </div>

  <div class="card" style="margin-bottom:24px">
    <div style="padding:20px 24px">
      <div class="sec-title" style="margin-bottom:16px">➕ Add New User</div>
      <form method="post" action="/create_user">
        <div class="g3">
          {_field_html("Username", '<input type="text" name="username" placeholder="e.g. jsmith" required>')}
          {_field_html("Password", '<input type="password" name="password" placeholder="Min. 6 characters" required minlength="6">')}
          {_field_html("Display Name", '<input type="text" name="display_name" placeholder="e.g. John Smith">')}
        </div>
        <div class="g3 mt">
          {_field_html("Role", _sel("role", ["user","director","pm","admin"], "user", "Select role…"))}
          {_field_html("Link to PM (if role = pm)", _sel("pm_name", get_project_managers(), "", "Select PM…"))}
          <div style="display:flex;align-items:flex-end">
            <button type="submit" class="btn-primary" style="width:100%;height:40px">+ Create User</button>
          </div>
        </div>

  <div style="background:#fff;border:1px solid {C['outline']};border-radius:10px;overflow:hidden">
    <div style="padding:12px 18px;border-bottom:1px solid {C['outline']};background:{C['surface_low']};display:flex;align-items:center;justify-content:space-between">
      <span style="font-size:13px;font-weight:700;color:{C['dark']}">All Users ({len(users)})</span>
      <span style="font-size:11px;color:{C['on_var']}">Role changes apply instantly • Click 🔑 to reset a password</span>
    </div>
    <div style="overflow-x:auto">
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead>
          <tr style="background:{C['surface_low']};border-bottom:1px solid {C['outline']}">
            <th style="text-align:left;padding:10px 16px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']};width:50px">ID</th>
            <th style="text-align:left;padding:10px 16px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Username</th>
            <th style="text-align:left;padding:10px 16px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Password Hash</th>
            <th style="text-align:left;padding:10px 16px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Display Name</th>
            <th style="text-align:left;padding:10px 16px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Role</th>
            <th style="text-align:left;padding:10px 16px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Linked PM</th>
            <th style="text-align:left;padding:10px 16px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Created At</th>
            <th style="text-align:right;padding:10px 16px;font-size:10px;font-weight:700;text-transform:uppercase;color:{C['on_var']}">Actions</th>
          </tr>
        </thead>
        <tbody>{rows_html}</tbody>
      </table>
    </div>
  </div>
</div></div>
</body></html>"""
    return html

def billing_collection(alert_html="", prefill=None, username="Admin") -> str:
    p     = prefill or {}
    today = date.today().isoformat()

    pm_list = get_project_managers()

    sel_pid        = p.get("project_id", "")
    sel_pm         = p.get("pm", "")
    sel_lob        = p.get("lob", "")   # retained for export links / legacy hidden fields
    pending_filter = p.get("pending_filter", "all")

    # ── A "pm" login only ever sees their own Billing & Collection entries —
    # ignore any other pm= value that may have been passed (e.g. a hand-edited
    # URL) and lock the PM selector to their own name. ─────────────────────────
    role         = (getattr(_req_ctx, "role", "") or "").lower()
    own_pm       = (getattr(_req_ctx, "pm_name", "") or "").strip()
    is_scoped_pm = bool(role == "pm" and own_pm)
    if is_scoped_pm:
        sel_pm = own_pm

    # ── PM dropdown (scoped PMs get a locked, read-only display instead) ──────
    if is_scoped_pm:
        pm_field_html = f"""
      <div class="field">
        <label>PM</label>
        <div style="background:{C['surface_low']};color:{C['dark']};font-size:13px;font-weight:700;
                    padding:9px 12px;border-radius:6px;border:1px solid {C['outline']};
                    display:inline-flex;align-items:center;gap:6px">👤 {html.escape(own_pm)}</div>
        <input type="hidden" name="pm" value="{html.escape(own_pm)}">
      </div>"""
    else:
        pm_opts = '<option value="">— Select PM —</option>' + "".join(
            f'<option value="{html.escape(str(pm))}" {"selected" if sel_pm == str(pm) else ""}>'
            f'{html.escape(str(pm))}</option>'
            for pm in pm_list
        )
        pm_field_html = f"""
      <div class="field">
        <label>PM</label>
        <select name="pm" onchange="this.form.submit()">{pm_opts}</select>
      </div>"""

    # ── Project dropdown: only populated once PM is chosen ────────────────────
    proj_rows = fetch_projects_for_pm(sel_pm) if sel_pm else []

    # ── Security: a scoped "pm" login must never be able to view another PM's
    # billing entries by hand-editing the project_id in the URL/form. Reject
    # any project_id that isn't actually one of this PM's own projects. ───────
    if is_scoped_pm and sel_pid:
        own_project_ids = {str(r[0]) for r in proj_rows}
        if str(sel_pid) not in own_project_ids:
            sel_pid = ""

    proj_opts_filter = '<option value="">— Select project —</option>' + "".join(
        f'<option value="{r[0]}" {"selected" if str(sel_pid) == str(r[0]) else ""}>'
        f'{r[1] or ""} | {r[2] or ""}</option>'
        for r in proj_rows
    )

    proj_field_html = ""
    if sel_pm:
        proj_field_html = f"""
      <div class="field">
        <label>Project Name</label>
        <select name="project_id" onchange="this.form.submit()">{proj_opts_filter}</select>
      </div>"""

    # ── Pending filter: only shown once project is chosen ─────────────────────
    pending_field_html = ""
    if sel_pm and sel_pid:
        pending_field_html = f"""
      <div class="field" style="max-width:180px">
        <label>Show</label>
        <select name="pending_filter" onchange="this.form.submit()">
          <option value="all"     {"selected" if pending_filter == "all"     else ""}>All entries</option>
          <option value="pending" {"selected" if pending_filter == "pending" else ""}>Pending only</option>
        </select>
      </div>"""

    # ── Data: only fetch when both PM and project are selected ────────────────
    show_data = bool(sel_pm and sel_pid)

    entries        = []
    filt_projects  = []
    rev_ov_sum     = itd_sum = phys_avg = eac_sum = fin_pct = 0
    total_billed   = total_collection = 0
    pending_count  = 0

    if show_data:
        entries = fetch_billing_entries_filtered(
            project_id=sel_pid or None,
            lob=sel_lob or None,
            pending_only=(pending_filter == "pending")
        )

        entry_pids   = {str(e.get("project_id")) for e in entries if e.get("project_id") not in (None, "")}
        all_projects = fetch_all_projects()
        all_projects = [pr for pr in all_projects if (pr.get("project_manager") or "").strip() == sel_pm.strip()]
        filt_projects = [pr for pr in all_projects if str(pr.get("id")) in entry_pids]
        if sel_pid:
            filt_projects = [pr for pr in filt_projects if str(pr.get("id")) == str(sel_pid)]

        rev_ov_sum = sum(_rev_ov(pr) for pr in filt_projects)
        itd_sum    = sum(_n(pr.get("itd_cost")) for pr in filt_projects)
        phys_avg   = (sum(_n(pr.get("physical_completion")) for pr in filt_projects) / len(filt_projects)) if filt_projects else 0
        eac_sum    = sum(_n(pr.get("rev_latest_eac") or pr.get("baseline_eac")) for pr in filt_projects)
        fin_pct    = (itd_sum / eac_sum * 100) if eac_sum else ((itd_sum / rev_ov_sum * 100) if rev_ov_sum else 0)

        total_billed     = sum(_n(e.get("billed_amt"))     for e in entries)
        total_collection = sum(_n(e.get("collection_amt")) for e in entries)
        pending_count    = sum(
            1 for e in entries
            if not ((e.get("collection_date") or "").strip() and _n(e.get("collection_amt")) > 0)
        )

    # ── Build table rows ──────────────────────────────────────────────────────
    rows_html = ""
    for e in entries:
        has_date   = bool((e.get("collection_date") or "").strip())
        has_amt    = _n(e.get("collection_amt")) > 0
        is_pending = not (has_date and has_amt)
        proj_label = f'{e.get("co_no","") or "—"}: {e.get("project_name","") or "—"}'

        if is_pending:
            coll_cell = (
                f'<td colspan="2" style="padding:10px 16px">'
                f'<form method="post" action="/record_collection" style="display:flex;gap:6px;align-items:center;flex-wrap:wrap">'
                f'<input type="hidden" name="entry_id"       value="{e["id"]}">'
                f'<input type="hidden" name="project_id"     value="{sel_pid}">'
                f'<input type="hidden" name="pm"             value="{sel_pm}">'
                f'<input type="hidden" name="pending_filter" value="{pending_filter}">'
                f'<input type="date"   name="collection_date" value="{e.get("collection_date","") or ""}" required class="bc-inline-inp" style="width:128px">'
                f'<input type="number" name="collection_amt"  value="{(_n(e.get("collection_amt")) or "") if _n(e.get("collection_amt")) > 0 else ""}" placeholder="Amount" step="0.01" required class="bc-inline-inp" style="width:100px">'
                f'<button type="submit" class="bc-record-btn">Record</button>'
                f'</form></td>'
            )
        else:
            coll_cell = (
                f'<td style="padding:12px 16px;color:{C["on_var"]}">{e.get("collection_date","")}</td>'
                f'<td style="padding:12px 16px;text-align:right;font-variant-numeric:tabular-nums;font-weight:700;color:{C["dark"]}">'
                f'₹ {_n(e.get("collection_amt")):,.0f}</td>'
            )

        row_cls = "is-pending" if is_pending else ""
        rows_html += f"""
        <tr class="{row_cls}">
          <td style="padding:12px 16px;font-weight:600;color:{C['dark']}">
            {proj_label}
            {'<span class="bc-badge bc-badge-pending">Pending</span>' if is_pending else ''}
          </td>
          <td style="padding:12px 16px;color:{C['on_var']}">{e.get('ra_no','') or '—'}</td>
          <td style="padding:12px 16px;color:{C['on_var']}">{e.get('billing_date','')}</td>
          <td style="padding:12px 16px;text-align:right;font-variant-numeric:tabular-nums;font-weight:700;color:{C['dark']}">₹ {_n(e.get('billed_amt')):,.0f}</td>
          {coll_cell}
          <td style="padding:12px 16px;text-align:right">
            {f'''<form method="post" action="/delete_billing_entry" style="display:inline">
              <input type="hidden" name="entry_id"       value="{e['id']}">
              <input type="hidden" name="project_id"     value="{sel_pid}">
              <input type="hidden" name="pm"             value="{sel_pm}">
              <input type="hidden" name="pending_filter" value="{pending_filter}">
              <button class="del-btn" onclick="return confirm(\'Delete this entry?\')" title="Delete entry">
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 7h16M9 7V4h6v3M6 7l1 14h10l1-14"/></svg>
              </button>
            </form>''' if is_pending else f'''<span class="bc-badge bc-badge-locked">
                <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="5" y="11" width="14" height="9" rx="1.5"/><path d="M8 11V7a4 4 0 1 1 8 0v4"/></svg>
                Locked
              </span>'''}
          </td>
        </tr>"""

    if show_data and not rows_html:
        rows_html = (
            f'<tr><td colspan="6" style="padding:40px;text-align:center;color:#aaa">'
            f'No billing entries match this filter. Use "+ Add Entry" to log one.</td></tr>'
        )

    # ── Side panel project options (pre-select current project) ───────────────
    if sel_pid:
        # Project already chosen — show as read-only label + hidden field in panel
        proj_panel_label = next(
            (f'{r[1] or ""} | {r[2] or ""}' for r in proj_rows if str(r[0]) == str(sel_pid)),
            "Selected project"
        )
        proj_panel_field = f"""
        <div class="sp-fld">
          <label>Project</label>
          <input type="hidden" name="project_id" value="{sel_pid}">
          <div style="padding:9px 12px;border:1px solid {C['outline']};border-radius:7px;
                      font-size:13px;color:{C['dark']};background:{C['surface_low']}">
            {proj_panel_label}
          </div>
        </div>"""
    else:
        proj_opts_panel = "".join(
            f'<option value="{r[0]}" {"selected" if str(sel_pid) == str(r[0]) else ""}>'
            f'{r[1] or ""} | {r[2] or ""}</option>'
            for r in proj_rows
        ) or '<option value="">— No projects yet —</option>'
        proj_panel_field = f"""
        <div class="sp-fld">
          <label>Project</label>
          <select name="project_id" required>
            <option value="">— Select project —</option>
            {proj_opts_panel}
          </select>
        </div>"""

    # ── Main content area ─────────────────────────────────────────────────────
    if not show_data:
        # Nothing selected yet — friendly prompt
        main_content = f"""
  <div style="padding:64px 40px;text-align:center;background:#fff;
              border:1px solid {C['outline']};border-radius:12px;
              box-shadow:0 1px 3px rgba(0,0,0,.04)">
    <svg width="40" height="40" viewBox="0 0 24 24" fill="none"
         stroke="{C['on_var']}" stroke-width="1.5" style="margin-bottom:16px;opacity:.5">
      <rect x="2" y="5" width="20" height="15" rx="2"/>
      <path d="M2 10h20M7 15h4"/>
    </svg>
    <p style="font-size:15px;font-weight:600;color:{C['dark']};margin:0 0 6px">
      {"Select a project to view billing entries" if sel_pm else "Start by selecting a PM"}
    </p>
    <p style="font-size:13px;color:{C['on_var']};margin:0">
      {"Use the Project filter above to load billing &amp; collection data." if sel_pm else "Choose a Project Manager from the filter above, then pick a project."}
    </p>
  </div>"""
    else:
        # Both LOB and project selected — show summary + table
        main_content = f"""
  <!-- Summary strip -->
  <div class="bc-stats">
    <div class="bc-stat"><div class="lbl">Order Value</div><div class="val">₹ {rev_ov_sum:,.0f}</div></div>
    <div class="bc-stat"><div class="lbl">Total Expenditure</div><div class="val">₹ {itd_sum:,.0f}</div></div>
    <div class="bc-stat"><div class="lbl">Total Billed (RA Bills)</div><div class="val">₹ {total_billed:,.0f}</div></div>
    <div class="bc-stat"><div class="lbl">Phy %</div><div class="val">{phys_avg:.1f}%</div></div>
    <div class="bc-stat"><div class="lbl">% Expenditure Over Order Value</div><div class="val">{fin_pct:.1f}%</div></div>
  </div>

  <!-- Entries table -->
  <div class="bc-card">
    <div class="bc-card-hdr">
      <h3>Billing Entries</h3>
      <div style="display:flex;align-items:center;gap:16px">
        <span class="count">{len(entries)} entries{f' &nbsp;·&nbsp; <b>{pending_count} pending</b>' if pending_count else ''}</span>
        <button type="button" class="btn-primary" onclick="openBCPanel()">+ Add Entry</button>
      </div>
    </div>
    <div style="overflow-x:auto">
      <table class="bc-tbl">
        <colgroup>
          <col style="width:25%">
          <col style="width:9%">
          <col style="width:14%">
          <col style="width:13%">
          <col style="width:15%">
          <col style="width:14%">
          <col style="width:10%">
        </colgroup>
        <thead>
          <tr>
            <th>Project</th>
            <th>RA No.</th>
            <th>Date of Billing</th>
            <th class="text-right">Billed Amt</th>
            <th>Date of Collection</th>
            <th class="text-right">Collection Amt</th>
            <th class="text-right">Action</th>
          </tr>
        </thead>
        <tbody>{rows_html}</tbody>
        <tfoot>
          <tr>
            <td colspan="3">Total</td>
            <td style="text-align:right;font-variant-numeric:tabular-nums">₹ {total_billed:,.0f}</td>
            <td></td>
            <td style="text-align:right;font-variant-numeric:tabular-nums">₹ {total_collection:,.0f}</td>
            <td></td>
          </tr>
        </tfoot>
      </table>
    </div>
  </div>"""

    # ── PM summary box: once a PM is chosen, show that PM's billing &
    #    collection totals across ALL of their projects (not just the one
    #    selected above), appended below the main Billing Entries card. ──
    if sel_pm:
        pm_summary_rows = fetch_billing_summary_for_pm(sel_pm)
        pm_total_billed     = sum(_n(r.get("total_billed"))     for r in pm_summary_rows)
        pm_total_collection = sum(_n(r.get("total_collection")) for r in pm_summary_rows)

        if pm_summary_rows:
            pm_rows_html = "".join(f"""
        <tr>
          <td style="padding:12px 16px;font-weight:600;color:{C['dark']}">{html.escape(str(r.get('co_no') or '—'))}: {html.escape(str(r.get('project_name') or '—'))}</td>
          <td style="padding:12px 16px;text-align:right;font-variant-numeric:tabular-nums;font-weight:700;color:{C['dark']}">₹ {_n(r.get('total_billed')):,.0f}</td>
          <td style="padding:12px 16px;text-align:right;font-variant-numeric:tabular-nums;font-weight:700;color:{C['dark']}">₹ {_n(r.get('total_collection')):,.0f}</td>
          <td style="padding:12px 16px;text-align:right;font-variant-numeric:tabular-nums;font-weight:700;color:{C['error'] if (_n(r.get('total_billed'))-_n(r.get('total_collection')))>0 else C['dark']}">₹ {(_n(r.get('total_billed'))-_n(r.get('total_collection'))):,.0f}</td>
        </tr>""" for r in pm_summary_rows)
        else:
            pm_rows_html = (
                f'<tr><td colspan="4" style="padding:32px;text-align:center;color:#aaa">'
                f'No projects found for this PM.</td></tr>'
            )

        main_content += f"""
  <div class="bc-card" style="margin-top:18px">
    <div class="bc-card-hdr">
      <h3>{html.escape(sel_pm)} — Billing &amp; Collection (All Projects)</h3>
      <span class="count">{len(pm_summary_rows)} project{'s' if len(pm_summary_rows) != 1 else ''}</span>
    </div>
    <div style="overflow-x:auto">
      <table class="bc-tbl" style="table-layout:auto">
        <thead>
          <tr>
            <th>Project</th>
            <th class="text-right">Total Billed</th>
            <th class="text-right">Total Collection</th>
            <th class="text-right">Pending</th>
          </tr>
        </thead>
        <tbody>{pm_rows_html}</tbody>
        <tfoot>
          <tr>
            <td>Total</td>
            <td style="text-align:right;font-variant-numeric:tabular-nums">₹ {pm_total_billed:,.0f}</td>
            <td style="text-align:right;font-variant-numeric:tabular-nums">₹ {pm_total_collection:,.0f}</td>
            <td style="text-align:right;font-variant-numeric:tabular-nums">₹ {(pm_total_billed - pm_total_collection):,.0f}</td>
          </tr>
        </tfoot>
      </table>
    </div>
  </div>"""

    # ── CSS ───────────────────────────────────────────────────────────────────
    css = f"""<style>
.bc-sel-bar{{background:#fff;border:1px solid {C['outline']};border-radius:12px;padding:18px 22px;display:flex;align-items:flex-end;gap:16px;flex-wrap:wrap;box-shadow:0 1px 3px rgba(0,0,0,.04)}}
.bc-sel-bar .field{{flex:1;min-width:180px}}
.bc-sel-bar .field label{{font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:{C['on_var']};display:block;margin-bottom:6px}}
.bc-sel-bar .field select{{width:100%;padding:9px 12px;border:1px solid {C['outline']};border-radius:7px;font-size:13px;background:#fff;outline:none;color:{C['dark']};transition:border-color .15s}}
.bc-sel-bar .field select:focus{{border-color:{C['dark']}}}

.bc-stats{{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;margin-bottom:18px}}
.bc-stat{{background:#fff;border:1px solid {C['outline']};border-left:3px solid {C['dark']};border-radius:10px;padding:14px 16px;box-shadow:0 1px 2px rgba(0,0,0,.03);transition:box-shadow .15s,transform .15s}}
.bc-stat:hover{{box-shadow:0 6px 16px rgba(0,0,0,.07);transform:translateY(-1px)}}
.bc-stat .lbl{{font-size:9px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['on_var']};margin-bottom:8px}}
.bc-stat .val{{font-size:19px;font-weight:800;color:{C['dark']};font-variant-numeric:tabular-nums}}

.bc-card{{background:#fff;border:1px solid {C['outline']};border-radius:12px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.04)}}
.bc-card-hdr{{display:flex;align-items:center;justify-content:space-between;padding:14px 18px;border-bottom:1px solid {C['outline']}}}
.bc-card-hdr h3{{font-size:14px;font-weight:700;color:{C['dark']}}}
.bc-card-hdr .count{{font-size:11px;color:{C['on_var']}}}
.bc-card-hdr .count b{{color:{C['dark']}}}

.bc-tbl{{width:100%;border-collapse:collapse;font-size:13px;table-layout:fixed}}
.bc-tbl td{{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.bc-tbl td:first-child{{white-space:normal}}
.bc-tbl thead th{{padding:11px 16px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;font-size:10.5px;color:{C['on_var']};text-align:left;background:{C['surface_low']};border-bottom:1.5px solid {C['outline']}}}
.bc-tbl thead th.text-right{{text-align:right}}
.bc-tbl tbody tr{{border-bottom:1px solid {C['outline']};transition:background .12s}}
.bc-tbl tbody tr:last-child{{border-bottom:none}}
.bc-tbl tbody tr:nth-child(even):not(.is-pending){{background:#fbfbfb}}
.bc-tbl tbody tr:hover{{background:{C['surface_low']}}}
.bc-tbl tbody tr.is-pending{{background:#fafaf8;box-shadow:inset 3px 0 0 {C['dark']}}}
.bc-tbl tbody tr.is-pending:hover{{background:#f3f3f0}}
.bc-tbl tfoot tr{{background:{C['surface_high']};border-top:2px solid {C['outline']}}}
.bc-tbl tfoot td{{padding:12px 16px;font-weight:700;color:{C['dark']}}}

.bc-badge{{display:inline-flex;align-items:center;gap:4px;margin-left:8px;padding:3px 9px;font-size:9px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;border-radius:20px;vertical-align:middle}}
.bc-badge-pending{{color:{C['dark']};background:#ececec;border:1px solid {C['outline']}}}
.bc-badge-locked{{color:{C['on_var']};background:#f1f1f1}}
.bc-inline-inp{{height:32px;border:1px solid {C['outline']};border-radius:6px;padding:0 9px;font-size:12px;color:{C['dark']}}}
.bc-record-btn{{height:32px;padding:0 14px;background:{C['dark']};color:#fff;border:none;border-radius:6px;font-size:11px;font-weight:700;cursor:pointer;transition:opacity .15s}}
.bc-record-btn:hover{{opacity:.85}}
.del-btn{{display:inline-flex;align-items:center;justify-content:center;width:28px;height:28px;background:none;border:none;border-radius:6px;cursor:pointer;color:{C['on_var']};opacity:.7;transition:opacity .15s,background .15s,color .15s}}
.del-btn:hover{{opacity:1;background:#f1f1f1;color:{C['dark']}}}

.side-panel{{position:fixed;inset-y:0;right:0;width:400px;background:#fff;border-left:1px solid {C['outline']};box-shadow:-4px 0 32px rgba(0,0,0,.10);z-index:300;transform:translateX(100%);transition:transform .3s ease;display:flex;flex-direction:column}}
.side-panel.open{{transform:translateX(0)}}
.sp-hdr{{padding:20px 24px;border-bottom:1px solid {C['outline']};background:{C['surface_low']};display:flex;align-items:center;justify-content:space-between}}
.sp-body{{flex:1;padding:24px;overflow-y:auto}}
.sp-section-lbl{{font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['on_var']};margin:20px 0 10px;padding-top:14px;border-top:1px solid {C['outline']}}}
.sp-section-lbl:first-child{{margin-top:0;padding-top:0;border-top:none}}
.sp-row{{display:flex;gap:10px}}
.sp-row .sp-fld{{flex:1}}
.sp-fld{{margin-bottom:14px}}
.sp-fld label{{font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:{C['on_var']};display:block;margin-bottom:6px}}
.sp-fld input,.sp-fld select{{width:100%;padding:9px 12px;border:1px solid {C['outline']};border-radius:7px;font-size:13px;background:#fff;outline:none;color:{C['dark']};transition:border-color .15s}}
.sp-fld input:focus,.sp-fld select:focus{{border-color:{C['dark']}}}
.sp-foot{{padding:16px 24px;border-top:1px solid {C['outline']};background:{C['surface_low']};display:flex;gap:10px}}
.overlay{{position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:299;display:none}}
.overlay.show{{display:block}}
</style>"""

    js = """<script>
function openBCPanel(){
  document.getElementById('bc-side-panel').classList.add('open');
  document.getElementById('bc-overlay').classList.add('show');
}
function closeBCPanel(){
  document.getElementById('bc-side-panel').classList.remove('open');
  document.getElementById('bc-overlay').classList.remove('show');
}
document.addEventListener('keydown', function(e){ if(e.key === 'Escape') closeBCPanel(); });
</script>"""

    # ── Assemble page ─────────────────────────────────────────────────────────
    page_html  = _html_head("Billing & Collection") + css + "<body>"
    page_html += _sidebar("billing_collection") + _topbar("BILLING & COLLECTION DETAILS", username)
    page_html += f"""
<div class="main"><div class="canvas" style="max-width:1280px">

  <!-- Page header -->
  <div style="margin-bottom:20px;display:flex;justify-content:space-between;align-items:flex-start">
    <div>
      <h2 style="font-size:24px;font-weight:800;color:{C['dark']};letter-spacing:-.01em">Billing &amp; Collection Details</h2>
      <p style="font-size:13px;color:#6b7280;margin-top:4px">Track RA-wise billing and the collections received across all projects.</p>
      <div style="width:42px;height:3px;background:{C['dark']};border-radius:2px;margin-top:12px;opacity:.85"></div>
    </div>
    {"" if not show_data else f'''<a href="/export_excel/billing?project_id={sel_pid}&pm={sel_pm}&pending_filter={pending_filter}"
       style="display:inline-flex;align-items:center;gap:6px;padding:8px 16px;
              background:{C["dark"]};color:#fff;border-radius:7px;font-size:12px;
              font-weight:700;text-decoration:none;letter-spacing:.04em">⬇ Export Excel</a>'''}
  </div>

  {alert_html}

  <!-- Filter bar -->
  <form method="post" action="/billing_collection_retrieve" style="margin-bottom:20px" id="bc-filter-form">
    <div class="bc-sel-bar">

      <!-- Step 1: PM — always visible (locked to own name for a "pm" login) -->
      {pm_field_html}

      <!-- Step 2: Project — only appears after PM is chosen -->
      {proj_field_html}

      <!-- Step 3: Pending filter — only appears after project is chosen -->
      {pending_field_html}

    </div>
  </form>

  <!-- Step 4 onwards: summary cards + entries table (or empty state) -->
  {main_content}

</div></div>

<!-- Add Entry side panel -->
<div class="side-panel" id="bc-side-panel">
  <div class="sp-hdr">
    <span style="font-size:16px;font-weight:700;color:{C['dark']}">Add Billing Entry</span>
    <button onclick="closeBCPanel()" type="button"
            style="background:none;border:none;font-size:20px;cursor:pointer;color:{C['on_var']}">✕</button>
  </div>
  <div class="sp-body">
    <form method="post" action="/save_billing_entry" id="bc-entry-form">
      <input type="hidden" name="pm"             value="{sel_pm}">
      <input type="hidden" name="pending_filter" value="{pending_filter}">

      <!-- Project: read-only display if already selected, dropdown otherwise -->
      {proj_panel_field}

      <div class="sp-section-lbl">Billing</div>
      <div class="sp-fld">
        <label>RA No.</label>
        <input type="text" name="ra_no" placeholder="e.g. RA-05">
      </div>
      <div class="sp-row">
        <div class="sp-fld">
          <label>Date of Billing</label>
          <input type="date" name="billing_date" value="{today}">
        </div>
        <div class="sp-fld">
          <label>Billed Amount (₹)</label>
          <input type="number" name="billed_amt" step="0.01" placeholder="0.00">
        </div>
      </div>

      <div class="sp-section-lbl">Collection <span style="font-weight:400;text-transform:none;letter-spacing:0">(optional)</span></div>
      <div class="sp-row">
        <div class="sp-fld">
          <label>Date of Collection</label>
          <input type="date" name="collection_date">
        </div>
        <div class="sp-fld">
          <label>Collection Amount (₹)</label>
          <input type="number" name="collection_amt" step="0.01" placeholder="0.00">
        </div>
      </div>
    </form>
  </div>
  <div class="sp-foot">
    <button onclick="closeBCPanel()" type="button"
            style="flex:1;padding:11px;border:1px solid {C['outline']};border-radius:7px;
                   background:#fff;color:{C['on_var']};font-size:11px;font-weight:700;
                   text-transform:uppercase;cursor:pointer">Cancel</button>
    <button onclick="document.getElementById('bc-entry-form').submit()" type="button"
            style="flex:1;padding:11px;background:{C['dark']};color:#fff;border:none;
                   border-radius:7px;font-size:11px;font-weight:700;
                   text-transform:uppercase;cursor:pointer">Save Entry</button>
  </div>
</div>
<div class="overlay" id="bc-overlay" onclick="closeBCPanel()"></div>

{js}
<script>
(function(){{
  var PAGE = 'billing_collection';
  var saved = MonarchFilters.init(PAGE);
  var url   = new URL(window.location.href);
  if (saved && !url.searchParams.get('project_id') && (saved.lob || saved.project_id)) {{
    var qs = [];
    if (saved.lob)            qs.push('lob='            + encodeURIComponent(saved.lob));
    if (saved.project_id)     qs.push('project_id='     + encodeURIComponent(saved.project_id));
    if (saved.pending_filter) qs.push('pending_filter=' + encodeURIComponent(saved.pending_filter));
    if (qs.length) {{ window.location.replace('/billing_collection?' + qs.join('&')); return; }}
  }}
  var form = document.getElementById('bc-filter-form');
  if (form) {{
    form.addEventListener('submit', function(){{
      var fd = new FormData(form);
      MonarchFilters.save(PAGE, {{
        lob: fd.get('lob')||'', project_id: fd.get('project_id')||'', pending_filter: fd.get('pending_filter')||'all'
      }});
    }});
  }}
}})();
</script>
</body></html>"""

    return page_html

# ═══════════════════════════════════════════════════════════════
#  PAGE 9 — WEEKLY PAYMENT BOOKING  (Screen 1 + Screen 2 approval workflow)
# ═══════════════════════════════════════════════════════════════

BOOKING_DOC_TYPES = [
    ("po_file",          "P.O."),
    ("invoice_file",     "Invoice"),
    ("quotation_file",   "Quotation/Other Docs"),
    ("vehicle_log_file", "Vehicle Usage Log"),
    ("vehicle_tag_file", "Vehicle Fuel Tag"),
    ("site_photo_file",  "Site Photos"),
]

PAYMENT_CATEGORIES = [
    "Vendor / Sub-contractor", "Material Supplier", "Consultancy Fees",
    "Utility / Statutory", "Travel & Conveyance", "Miscellaneous",
]

BOOKING_LOCATIONS = [
    "Business Development", "Monarch House", "KK", "Morya", "Cybernex",
    "Nagpur Office", "Delhi Office", "Nashik Office", "Gandhinagar Office",
    "Patna Office", "Mumbai Office", "Others",
]

UPLOAD_DIR = os.path.join("uploads", "payment_bookings")


def _parse_multipart_full(handler) -> tuple[dict, dict]:
    """Parses multipart/form-data into (text_fields, files) where files maps
    field name -> list of (filename, bytes) for every file part that has content."""
    import email as _email, email.policy as _email_policy
    content_type = handler.headers.get("Content-Type", "")
    n = int(handler.headers.get("Content-Length", 0))
    raw = handler.rfile.read(n)
    fields, files = {}, {}
    if "multipart/form-data" not in content_type:
        return fields, files
    full_msg_bytes = f"Content-Type: {content_type}\r\n\r\n".encode() + raw
    try:
        msg = _email.message_from_bytes(full_msg_bytes, policy=_email_policy.compat32)
    except Exception:
        return fields, files
    for part in msg.walk():
        cd = part.get("Content-Disposition", "")
        if not cd:
            continue
        name, filename = None, None
        for seg in cd.split(";"):
            seg = seg.strip()
            if seg.startswith("name="):
                name = seg.split("=", 1)[1].strip().strip('"\'')
            elif seg.startswith("filename="):
                filename = seg.split("=", 1)[1].strip().strip('"\'')
        if name is None:
            continue
        payload = part.get_payload(decode=True)
        if filename:
            if payload:
                files.setdefault(name, []).append((filename, payload))
        else:
            if payload is not None:
                fields[name] = payload.decode("utf-8", errors="replace")
    return fields, files


def get_booking_period(ref: date = None) -> tuple[str, str]:
    """Current Mon–Sun weekly booking window."""
    ref = ref or date.today()
    start = ref - timedelta(days=ref.weekday())
    end = start + timedelta(days=6)
    return start.isoformat(), end.isoformat()


def eac_guard_status(proj: dict) -> tuple[bool, str]:
    """Returns (blocked, message). Booking is blocked when EAC < ITD cost or EAC = 0."""
    if not proj:
        return False, ""
    eac = _n(proj.get("rev_latest_eac")) or _n(proj.get("baseline_eac"))
    itd = _n(proj.get("itd_cost"))
    if eac == 0:
        return True, "⚠ EAC has not been set for this project. Please get the EAC revised before booking further payments."
    if eac < itd:
        return True, "⚠ EAC is lower than the ITD cost for this project. Please get the EAC revised before booking further payments."
    return False, ""


def fetch_bookable_projects_for_pm(pm_name: str) -> list[dict]:
    try:
        with ENGINE.connect() as conn:
            rows = _rows_to_dicts(conn, "SELECT * FROM projects ORDER BY project_name")
        pm_l = (pm_name or "").strip().lower()
        if not pm_l:
            return rows
        return [r for r in rows if (r.get("project_manager") or "").strip().lower() == pm_l]
    except Exception as e:
        print(f"[fetch_bookable_projects_for_pm] {e}")
        return []


def _save_booking_file(booking_id: int, doc_type: str, filename: str, data: bytes):
    try:
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        safe_name = re.sub(r'[^A-Za-z0-9._-]+', '_', filename or "file")
        stored_name = f"{booking_id}_{secrets.token_hex(4)}_{safe_name}"
        path = os.path.join(UPLOAD_DIR, stored_name)
        with open(path, "wb") as f:
            f.write(data)
        with ENGINE.begin() as conn:
            conn.execute(text("""
                INSERT INTO payment_booking_docs (booking_id, doc_type, file_name, file_path, uploaded_at)
                VALUES (:b,:t,:n,:p,:u)
            """), {"b": booking_id, "t": doc_type, "n": filename, "p": path, "u": datetime.now().isoformat()})
    except Exception as e:
        print(f"[_save_booking_file] {e}")


def save_payment_booking(fields: dict, files: dict, username: str, role: str, pm_name: str) -> tuple[bool, str, str]:
    booking_id_raw = (fields.get("booking_id") or "").strip()
    is_edit = booking_id_raw.isdigit()
    edit_bid = int(booking_id_raw) if is_edit else None

    if is_edit:
        try:
            with ENGINE.connect() as conn:
                row = conn.execute(text(
                    "SELECT pm_username, overall_status, booking_uid FROM payment_bookings WHERE id=:id"
                ), {"id": edit_bid}).fetchone()
        except Exception as e:
            return False, f"Could not load booking: {e}", ""
        if not row:
            return False, "Booking not found.", ""
        if (row[0] or "").strip().lower() != (username or "").strip().lower():
            return False, "You can only edit your own bookings.", ""
        if (row[1] or "").lower() != "disapproved":
            return False, "Only a disapproved booking can be edited and resent.", ""
        existing_uid = row[2]

    is_indirect = fields.get("is_indirect") == "1"
    book_mode = (fields.get("book_mode") or "project").strip().lower()
    project_id = None
    indirect_dept = ""
    eac_bypassed = False
    location = (fields.get("location") or "").strip()

    if is_indirect:
        if not is_head_role(role):
            return False, "Only Heads of Department can book Indirect Expenses.", ""
        indirect_dept = (fields.get("indirect_dept") or "").strip()
        if not indirect_dept:
            return False, "Please specify the department for Indirect Expenses.", ""
        project_label = f"Indirect Expenses — {indirect_dept}"
        if not location:
            return False, "Please select a location.", ""
    elif book_mode == "location":
        # Booking against a Location only — no project selected.
        if not location:
            return False, "Please select a location.", ""
        project_label = location
    else:
        pid_raw = (fields.get("project_id") or "").strip()
        if not pid_raw.isdigit():
            return False, "Please select a project.", ""
        project_id = int(pid_raw)
        proj = fetch_project_by_id(project_id)
        if not proj:
            return False, "Selected project could not be found.", ""
        if (pm_name and not is_procurement_role(role)
                and (proj.get("project_manager") or "").strip().lower() != pm_name.strip().lower()):
            return False, "You can only book payments against projects assigned to you.", ""
        blocked, guard_msg = eac_guard_status(proj)
        if blocked:
            if fields.get("eac_bypass") == "1":
                eac_bypassed = True
            else:
                return False, guard_msg, ""
        project_label = f'{proj.get("co_no","") or ""}: {proj.get("project_name","")}'.strip(": ")

    payee_name = (fields.get("payee_name") or "").strip()
    payment_category = (fields.get("payment_category") or "").strip()
    requested_amount = _n(fields.get("requested_amount"))
    invoice_date = (fields.get("invoice_date") or "").strip()
    justification = (fields.get("justification") or "").strip()

    if not (payee_name and payment_category and requested_amount and invoice_date and justification):
        return False, "Please fill all the required fields before submitting.", ""

    if not is_indirect and eac_bypassed:
        justification = f"[EAC CHECK BYPASSED] {justification}"

    period_start, period_end = get_booking_period()

    if is_edit:
        try:
            with ENGINE.begin() as conn:
                conn.execute(text("""
                    UPDATE payment_bookings SET
                        project_id=:pid, is_indirect=:ind, indirect_dept=:dept, location=:loc, project_label=:lbl,
                        payee_name=:payee, payment_category=:cat, requested_amount=:amt,
                        invoice_date=:idate, justification=:just,
                        period_start=:ps, period_end=:pe,
                        overall_status=NULL,
                        procurement_status=NULL, procurement_note=NULL, procurement_by=NULL, procurement_at=NULL,
                        accounts_status=NULL,    accounts_note=NULL,    accounts_by=NULL,    accounts_at=NULL,
                        director_status=NULL,    director_note=NULL,    director_by=NULL,    director_at=NULL,
                        archived=FALSE
                    WHERE id=:id
                """), {
                    "pid": project_id, "ind": is_indirect, "dept": indirect_dept, "loc": location, "lbl": project_label,
                    "payee": payee_name, "cat": payment_category, "amt": requested_amount,
                    "idate": invoice_date, "just": justification,
                    "ps": period_start, "pe": period_end, "id": edit_bid,
                })
        except Exception as e:
            print(f"[save_payment_booking:edit] {e}")
            return False, f"Could not update booking: {e}", ""
        new_id = edit_bid
        booking_uid = existing_uid
    else:
        try:
            with ENGINE.begin() as conn:
                result = conn.execute(text("""
                    INSERT INTO payment_bookings
                        (project_id, is_indirect, indirect_dept, location, project_label, pm_username, pm_name,
                         payee_name, payment_category, requested_amount, invoice_date, justification,
                         period_start, period_end, digitally_signed_by, created_at)
                    VALUES (:pid,:ind,:dept,:loc,:lbl,:un,:pn,:payee,:cat,:amt,:idate,:just,:ps,:pe,:sig,:ca)
                    RETURNING id
                """), {
                    "pid": project_id, "ind": is_indirect, "dept": indirect_dept, "loc": location, "lbl": project_label,
                    "un": username, "pn": pm_name or username,
                    "payee": payee_name, "cat": payment_category, "amt": requested_amount,
                    "idate": invoice_date, "just": justification,
                    "ps": period_start, "pe": period_end, "sig": username,
                    "ca": datetime.now().isoformat(),
                })
                new_id = result.scalar()
                booking_uid = f"PB-{date.today().year}-{new_id:05d}"
                conn.execute(text("UPDATE payment_bookings SET booking_uid=:u WHERE id=:id"),
                             {"u": booking_uid, "id": new_id})
        except Exception as e:
            print(f"[save_payment_booking] {e}")
            return False, f"Could not save booking: {e}", ""

    for field_name, doc_type in BOOKING_DOC_TYPES:
        for filename, data in files.get(field_name, []):
            _save_booking_file(new_id, doc_type, filename, data)

    msg = (f"Payment booking {booking_uid} updated and resent for approval." if is_edit
           else f"Payment booking {booking_uid} submitted for approval.")
    return True, msg, booking_uid


def fetch_bookings_for_pm(username: str) -> list[dict]:
    try:
        with ENGINE.connect() as conn:
            return _rows_to_dicts(conn,
                "SELECT * FROM payment_bookings WHERE pm_username=:u ORDER BY id DESC", {"u": username})
    except Exception as e:
        print(f"[fetch_bookings_for_pm] {e}")
        return []


def fetch_all_bookings(include_archived: bool = True) -> list[dict]:
    try:
        with ENGINE.connect() as conn:
            sql = "SELECT * FROM payment_bookings"
            if not include_archived:
                sql += " WHERE archived = FALSE"
            sql += " ORDER BY id DESC"
            return _rows_to_dicts(conn, sql)
    except Exception as e:
        print(f"[fetch_all_bookings] {e}")
        return []


def fetch_booking_detail(booking_id: int) -> dict:
    try:
        with ENGINE.connect() as conn:
            rows = _rows_to_dicts(conn, "SELECT * FROM payment_bookings WHERE id=:id", {"id": booking_id})
            b = rows[0] if rows else {}
            if b:
                b["docs"] = _rows_to_dicts(conn,
                    "SELECT * FROM payment_booking_docs WHERE booking_id=:id ORDER BY id", {"id": booking_id})
            return b
    except Exception as e:
        print(f"[fetch_booking_detail] {e}")
        return {}


def procurement_decide(booking_id: int, decision: str, note: str, actor: str) -> bool:
    status = "approved" if decision == "approve" else "disapproved"
    try:
        with ENGINE.begin() as conn:
            conn.execute(text("""
                UPDATE payment_bookings SET procurement_status=:s, procurement_note=:n,
                       procurement_by=:a, procurement_at=:t WHERE id=:id
            """), {"s": status, "n": note, "a": actor, "t": datetime.now().isoformat(), "id": booking_id})
            if status == "disapproved":
                conn.execute(text("UPDATE payment_bookings SET overall_status='disapproved' WHERE id=:id"),
                             {"id": booking_id})
        return True
    except Exception as e:
        print(f"[procurement_decide] {e}")
        return False


def accounts_decide(booking_id: int, decision: str, note: str, actor: str) -> bool:
    status = "approved" if decision == "approve" else "disapproved"
    try:
        with ENGINE.begin() as conn:
            conn.execute(text("""
                UPDATE payment_bookings SET accounts_status=:s, accounts_note=:n,
                       accounts_by=:a, accounts_at=:t WHERE id=:id
            """), {"s": status, "n": note, "a": actor, "t": datetime.now().isoformat(), "id": booking_id})
            if status == "disapproved":
                conn.execute(text("UPDATE payment_bookings SET overall_status='disapproved' WHERE id=:id"),
                             {"id": booking_id})
        return True
    except Exception as e:
        print(f"[accounts_decide] {e}")
        return False


def director_decide(booking_id: int, decision: str, note: str, actor: str) -> bool:
    status = "approved" if decision == "approve" else "disapproved"
    try:
        with ENGINE.begin() as conn:
            conn.execute(text("""
                UPDATE payment_bookings SET director_status=:s, director_note=:n,
                       director_by=:a, director_at=:t, overall_status=:ov, archived=:arc
                WHERE id=:id
            """), {
                "s": status, "n": note, "a": actor, "t": datetime.now().isoformat(),
                "ov": status, "arc": (status == "approved"), "id": booking_id,
            })
        return True
    except Exception as e:
        print(f"[director_decide] {e}")
        return False


def fetch_booking_stats() -> dict:
    try:
        with ENGINE.connect() as conn:
            total    = conn.execute(text("SELECT COUNT(*) FROM payment_bookings")).scalar() or 0
            approved = conn.execute(text("SELECT COUNT(*) FROM payment_bookings WHERE overall_status='approved'")).scalar() or 0
            pending  = conn.execute(text(
                "SELECT COUNT(*) FROM payment_bookings "
                "WHERE overall_status IS NULL OR overall_status NOT IN ('approved','disapproved')"
            )).scalar() or 0
        return {"total": total, "approved": approved, "pending": pending}
    except Exception as e:
        print(f"[fetch_booking_stats] {e}")
        return {"total": 0, "approved": 0, "pending": 0}


def export_payment_bookings_excel_bytes() -> bytes:
    """3 sub-sheets: Approved / Pending / Disapproved, each with all 3 authorities' comments."""
    from openpyxl.styles import Font, PatternFill, Alignment
    bookings = fetch_all_bookings(include_archived=True)
    wb = Workbook()
    wb.remove(wb.active)
    hdr_fill = PatternFill("solid", fgColor="000E24")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    headers = ["Booking ID", "Date", "Project / Indirect Dept", "Payee", "Category",
               "Amount (Rs)", "PM / Head", "Procurement Note", "Accounts Note", "Director Note", "Status"]
    groups = [
        ("Approved",    [b for b in bookings if b.get("overall_status") == "approved"]),
        ("Pending",     [b for b in bookings if b.get("overall_status") not in ("approved", "disapproved")]),
        ("Disapproved", [b for b in bookings if b.get("overall_status") == "disapproved"]),
    ]
    for sheet_name, rows in groups:
        ws = wb.create_sheet(sheet_name)
        ws.append(headers)
        for cell in ws[1]:
            cell.fill = hdr_fill; cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for b in rows:
            ws.append([
                b.get("booking_uid", ""), str(b.get("created_at", ""))[:10], b.get("project_label", ""),
                b.get("payee_name", ""), b.get("payment_category", ""), _n(b.get("requested_amount")),
                b.get("pm_name", ""), b.get("procurement_note", "") or "",
                b.get("accounts_note", "") or "", b.get("director_note", "") or "",
                (b.get("overall_status", "pending") or "pending").upper(),
            ])
        for idx, w in enumerate([14, 12, 30, 18, 18, 14, 16, 26, 26, 26, 12], start=1):
            ws.column_dimensions[ws.cell(1, idx).column_letter].width = w
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf.read()


def stamp_booking_pdf_bytes(booking_id: int) -> bytes:
    """Generates a stamped approval certificate for a booking — 'DIGITALLY APPROVED
    FOR PAYMENT' plus the Director's note printed on it once fully approved."""
    b = fetch_booking_detail(booking_id)
    styles = getSampleStyleSheet()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=40, bottomMargin=40, leftMargin=40, rightMargin=40)
    elements = [
        Paragraph("Monarch Surveyors &amp; Engineering Consultants Ltd.", styles["Title"]),
        Spacer(1, 6),
        Paragraph(f"Payment Booking — {b.get('booking_uid','')}", styles["Heading2"]),
        Spacer(1, 10),
    ]
    data = [
        ["Project / Dept",   b.get("project_label", "")],
        ["Payee",            b.get("payee_name", "")],
        ["Category",         b.get("payment_category", "")],
        ["Requested Amount", f'Rs {_n(b.get("requested_amount")):,.2f}'],
        ["Invoice Date",     b.get("invoice_date", "")],
        ["Booked By",        b.get("digitally_signed_by", "")],
        ["Procurement",      f'{b.get("procurement_status","")} — {b.get("procurement_note","") or ""}'],
        ["Accounts",         f'{b.get("accounts_status","")} — {b.get("accounts_note","") or ""}'],
        ["Director",         f'{b.get("director_status","")} — {b.get("director_note","") or ""}'],
    ]
    tbl = Table(data, colWidths=[140, 360])
    tbl.setStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, grey),
        ("BACKGROUND", (0, 0), (0, -1), Color(0.94, 0.96, 0.98)),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ])
    elements.append(tbl)
    elements.append(Spacer(1, 24))
    if b.get("overall_status") == "approved":
        elements.append(Paragraph('<font color="green"><b>DIGITALLY APPROVED FOR PAYMENT</b></font>', styles["Heading3"]))
        if b.get("director_note"):
            elements.append(Spacer(1, 6))
            elements.append(Paragraph(f'Director\'s note: {html.escape(b.get("director_note",""))}', styles["Normal"]))
    docs = b.get("docs") or []
    if docs:
        elements.append(Spacer(1, 16))
        elements.append(Paragraph("Attached Documents", styles["Heading4"]))
        for d in docs:
            elements.append(Paragraph(f'- {d.get("doc_type","")}: {html.escape(d.get("file_name","") or "")}', styles["Normal"]))
    doc.build(elements)
    buf.seek(0)
    return buf.read()


def weekly_payment_booking(alert_html="", username="Admin", role="pm", pm_name="",
                            sel_project_id="", is_indirect=False, edit_booking: dict = None) -> str:
    period_start, period_end = get_booking_period()
    edit_id_val = edit_booking.get("id", "") if edit_booking else ""
    ef = edit_booking or {}
    def _ev(key):
        return html.escape(str(ef.get(key) or ""))
    # Both PM and Head roles must only ever see projects they are assigned to
    # as Project Manager — never the full project list. Only elevated/admin
    # roles (procurement, accounts, director, admin) fall through to "" and
    # see everything, which is intentional for their oversight screens.
    projects = fetch_bookable_projects_for_pm(pm_name if is_booking_role(role) else "")
    proj = fetch_project_by_id(int(sel_project_id)) if str(sel_project_id).isdigit() else {}
    guard_blocked, guard_msg = eac_guard_status(proj) if proj else (False, "")

    proj_opts = '<option value="">— Select project —</option>' + "".join(
        f'<option value="{p["id"]}" {"selected" if str(sel_project_id) == str(p["id"]) else ""}>'
        f'{p.get("co_no","") or ""} — {p.get("project_name","")}</option>'
        for p in projects
    )

    info_card = ""
    if proj:
        order_val = _n(proj.get('order_value'))
        itd_val   = _n(proj.get('itd_cost'))
        eac = _n(proj.get("rev_latest_eac")) or _n(proj.get("baseline_eac"))
        pct = min(100, round((itd_val / eac) * 100)) if eac else 0
        pct_cls = "ok" if pct < 80 else ("warn" if pct < 100 else "err")
        info_card = f"""
        <div class="wpb-proj-snap">
          <div class="wpb-proj-snap-hdr">
            <span class="wpb-proj-snap-name">🏗 {proj.get('project_name','')}</span>
            <span class="wpb-proj-snap-co">{proj.get('co_no','') or ''}</span>
          </div>
          <div class="wpb-info">
            <div class="wpb-info-row"><span>Order Value</span><b>₹ {order_val:,.0f}</b></div>
            <div class="wpb-info-row"><span>ITD Cost</span><b>₹ {itd_val:,.0f}</b></div>
            <div class="wpb-info-row"><span>EAC</span><b>₹ {eac:,.0f}</b></div>
          </div>
          <div class="wpb-util-bar"><div class="wpb-util-fill wpb-util-{pct_cls}" style="width:{pct}%"></div></div>
          <div class="wpb-util-lbl">{pct}% of EAC spent to date</div>
        </div>"""

    warn_html = ""
    if guard_blocked:
        warn_html = f"""
        <div class="alert err" id="wpb-eac-alert">
          <div>{guard_msg}</div>
          <div style="margin-top:10px">
            <button type="button" id="wpb-bypass-btn" onclick="wpbBypass()"
              style="padding:7px 14px;font-size:12.5px;font-weight:700;border-radius:7px;
                     border:1px solid {C['error']};background:#fff;color:{C['error']};cursor:pointer">
              Bypass
            </button>
          </div>
        </div>
        """

    head_toggle = ""
    if is_head_role(role):
        head_toggle = f"""
        <div class="sp-fld">
          <label class="wpb-switch-lbl">
            <span class="wpb-switch">
              <input type="checkbox" name="is_indirect" value="1" id="wpb-indirect"
                     {"checked" if is_indirect else ""}
                     onchange="document.getElementById('wpb-dept-fld').style.display=this.checked?'block':'none';
                               document.getElementById('wpb-mode-row').style.display=this.checked?'none':'flex';
                               if(this.checked){{document.getElementById('wpb-location-select').required=false;}}
                               else{{wpbSetBookMode('project');}}">
              <span class="wpb-switch-track"><span class="wpb-switch-thumb"></span></span>
            </span>
            <span>This is an Indirect Expense (Admin / HR / IT etc.)</span>
          </label>
        </div>
        <div class="sp-fld" id="wpb-dept-fld" style="display:{'block' if is_indirect else 'none'}">
          <label>Department</label>
          <input type="text" name="indirect_dept" value="{_ev('indirect_dept')}" placeholder="e.g. Admin, HR, IT">
        </div>"""

    cat_opts = "".join(f'<option value="{c}">{c}</option>' for c in PAYMENT_CATEGORIES)
    DOC_ICONS = {
        "po_file": "📄", "invoice_file": "🧾", "quotation_file": "📎",
        "vehicle_log_file": "🚚", "vehicle_tag_file": "⛽", "site_photo_file": "📷",
    }
    doc_inputs = "".join(
        f'<label class="wpb-doc" for="wpb-file-{fname}">'
        f'<span class="wpb-doc-ico">{DOC_ICONS.get(fname,"📁")}</span>'
        f'<span class="wpb-doc-lbl">{label}</span>'
        f'<span class="wpb-doc-name" id="wpb-name-{fname}">Click to attach</span>'
        f'<input type="file" name="{fname}" id="wpb-file-{fname}" multiple '
        f'onchange="wpbShowFiles(this,\'wpb-name-{fname}\')"></label>'
        for fname, label in BOOKING_DOC_TYPES
    )

    my_bookings = fetch_bookings_for_pm(username)
    n_pending  = sum(1 for b in my_bookings if (b.get("overall_status") or "pending").lower() not in ("approved", "disapproved"))
    n_approved = sum(1 for b in my_bookings if (b.get("overall_status") or "").lower() == "approved")
    amt_period = sum(_n(b.get("requested_amount")) for b in my_bookings
                      if str(b.get("period_start","")) == period_start)

    STAGE_ICONS = {"pending": "⏳", "approved": "✓", "disapproved": "✕"}
    rows_html = ""
    for b in my_bookings:
        status = (b.get("overall_status") or "pending").lower()
        badge_cls = {"approved": "ok", "disapproved": "err"}.get(status, "pend")

        def _mini_stage(label, skey):
            s = (b.get(skey) or "pending").lower()
            cls = {"approved": "ok", "disapproved": "err"}.get(s, "pend")
            return f'<span class="wpb-ministage wpb-ministage-{cls}">{STAGE_ICONS.get(s,"⏳")} {label}</span>'

        rows_html += f"""
        <tr onclick="toggleNotes({b['id']})" class="wpb-row{' is-pending' if badge_cls=='pend' else ''}">
          <td style="font-weight:700;color:{C['dark']}">{b.get('booking_uid','')}</td>
          <td>{b.get('project_label','')}</td>
          <td>{b.get('location','') or '—'}</td>
          <td>{b.get('payee_name','')}</td>
          <td style="text-align:right;font-weight:600">₹ {_n(b.get('requested_amount')):,.0f}</td>
          <td>{str(b.get('created_at',''))[:10]}</td>
          <td><span class="wpb-badge wpb-badge-{badge_cls}">{status.title()}</span></td>
          <td>{f'''<a href="/weekly_booking?edit_id={b['id']}" onclick="event.stopPropagation()"
                  style="display:inline-block;padding:5px 12px;font-size:11.5px;font-weight:700;border-radius:6px;
                  background:{C['dark']};color:#fff;text-decoration:none;white-space:nowrap">↻ Update</a>''' if status == "disapproved" else ""}</td>
        </tr>
        <tr id="notes-{b['id']}" class="wpb-notes-row" style="display:none">
          <td colspan="8">
            <div class="wpb-notes-wrap">
              {_mini_stage("Procurement", "procurement_status")}
              {_mini_stage("Accounts", "accounts_status")}
              {_mini_stage("Director", "director_status")}
              <div class="wpb-notes-text">
                {f'<div><b>Procurement:</b> {b.get("procurement_note")}</div>' if b.get("procurement_note") else ""}
                {f'<div><b>Accounts:</b> {b.get("accounts_note")}</div>' if b.get("accounts_note") else ""}
                {f'<div><b>Director:</b> {b.get("director_note")}</div>' if b.get("director_note") else ""}
              </div>
            </div>
          </td>
        </tr>"""
    if not rows_html:
        rows_html = ('<tr><td colspan="7"><div class="wpb-booked-empty">'
                     '<div class="ico">🗂️</div>No payment bookings yet — '
                     'use the form above to submit your first one.</div></td></tr>')

    css = f"""<style>
.wpb-period-row{{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:20px}}
.wpb-period{{background:linear-gradient(135deg,{C['dark']},{C['mid']});color:#fff;border-radius:10px;padding:12px 18px;font-size:12px;font-weight:700;
  display:inline-flex;gap:10px;align-items:center;box-shadow:0 2px 8px rgba(0,14,36,.15)}}
.wpb-mini-stat{{background:#fff;border:1px solid {C['outline']};border-radius:10px;padding:9px 16px;font-size:11px;color:{C['on_var']};display:inline-flex;gap:8px;align-items:center}}
.wpb-mini-stat b{{color:{C['dark']};font-size:14px}}
.wpb-card{{background:#fff;border:1px solid {C['outline']};border-radius:14px;padding:24px;box-shadow:0 1px 3px rgba(0,0,0,.04);margin-bottom:20px}}
.wpb-proj-snap{{background:{C['surface_low']};border:1px solid {C['outline']};border-radius:10px;padding:14px 16px;margin:14px 0 18px}}
.wpb-proj-snap-hdr{{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:10px}}
.wpb-proj-snap-name{{font-weight:700;color:{C['dark']};font-size:13px}}
.wpb-proj-snap-co{{font-size:11px;color:{C['on_var']}}}
.wpb-info{{display:flex;gap:22px;margin:0 0 12px;flex-wrap:wrap}}
.wpb-info-row{{background:#fff;border:1px solid {C['outline']};border-radius:8px;padding:8px 14px;font-size:12px;display:flex;gap:8px;align-items:center}}
.wpb-info-row span{{color:{C['on_var']}}}
.wpb-info-row b{{color:{C['dark']}}}
.wpb-util-bar{{height:6px;border-radius:4px;background:{C['surface_high']};overflow:hidden}}
.wpb-util-fill{{height:100%;border-radius:4px;transition:width .3s}}
.wpb-util-ok{{background:#2fa858}}
.wpb-util-warn{{background:{C['secondary']}}}
.wpb-util-err{{background:{C['error']}}}
.wpb-util-lbl{{font-size:10.5px;color:{C['on_var']};margin-top:6px}}
.wpb-switch-lbl{{display:flex;align-items:center;gap:10px;cursor:pointer;font-size:13px;color:{C['on_surface']};font-weight:500;text-transform:none;letter-spacing:normal}}
.wpb-switch{{position:relative;display:inline-block;width:38px;height:22px;flex-shrink:0}}
.wpb-switch input{{opacity:0;width:0;height:0;position:absolute}}
.wpb-switch-track{{position:absolute;inset:0;background:{C['outline']};border-radius:22px;transition:background .18s;display:block}}
.wpb-switch-thumb{{position:absolute;top:2px;left:2px;width:18px;height:18px;background:#fff;border-radius:50%;transition:transform .18s;box-shadow:0 1px 2px rgba(0,0,0,.25)}}
.wpb-switch input:checked + .wpb-switch-track{{background:{C['secondary']}}}
.wpb-switch input:checked + .wpb-switch-track .wpb-switch-thumb{{transform:translateX(16px)}}
.wpb-docs{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-top:6px}}
.wpb-doc{{display:flex;flex-direction:column;align-items:center;text-align:center;gap:4px;padding:16px 10px;
  border:1.5px dashed {C['outline']};border-radius:10px;cursor:pointer;transition:border-color .15s,background .15s;background:#fbfbfc}}
.wpb-doc:hover{{border-color:{C['secondary']};background:#fdf8ee}}
.wpb-doc-ico{{font-size:20px}}
.wpb-doc-lbl{{font-size:10.5px;font-weight:700;letter-spacing:.03em;color:{C['on_surface']}}}
.wpb-doc-name{{font-size:10px;color:{C['on_var']};max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.wpb-doc input[type=file]{{display:none}}
.wpb-badge{{padding:3px 10px;border-radius:20px;font-size:10px;font-weight:700;text-transform:uppercase;display:inline-block}}
.wpb-badge-ok{{background:#e3f5e6;color:#1a7a34}}
.wpb-badge-err{{background:#fbe4e4;color:{C['error']}}}
.wpb-badge-pend{{background:#fdf3da;color:{C['secondary']}}}
.wpb-row{{cursor:pointer}}
.wpb-row.is-pending{{box-shadow:inset 3px 0 0 {C['secondary']}}}
.wpb-notes-row td{{padding:0 !important;background:{C['surface_low']}}}
.wpb-notes-wrap{{padding:12px 16px;font-size:12px;color:{C['on_var']}}}
.wpb-notes-text{{margin-top:8px;display:flex;flex-direction:column;gap:3px}}
.wpb-ministage{{display:inline-flex;align-items:center;gap:4px;font-size:10.5px;font-weight:700;padding:3px 9px;border-radius:20px;margin-right:6px}}
.wpb-ministage-ok{{background:#e3f5e6;color:#1a7a34}}
/* ── List of Payments Booked: scoped card + table so it always renders
     correctly regardless of any other .bc-card / .bc-tbl rules on the page ── */
.wpb-booked-card{{background:#fff;border:1px solid {C['outline']};border-radius:14px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.04)}}
.wpb-booked-hdr{{display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid {C['outline']}}}
.wpb-booked-hdr h3{{font-size:15px;font-weight:800;color:{C['dark']};margin:0}}
.wpb-booked-count{{font-size:11.5px;font-weight:600;color:{C['on_var']};background:{C['surface_low']};padding:4px 12px;border-radius:20px;flex-shrink:0}}
.wpb-booked-table-wrap{{overflow-x:auto}}
.wpb-booked-table{{width:100%;border-collapse:collapse;font-size:13px}}
.wpb-booked-table thead th{{padding:12px 18px;font-weight:700;font-size:10.5px;letter-spacing:.06em;text-transform:uppercase;
  color:{C['on_var']};text-align:left;background:{C['surface_low']};border-bottom:1.5px solid {C['outline']};white-space:nowrap}}
.wpb-booked-table thead th.text-right{{text-align:right}}
.wpb-booked-table tbody tr.wpb-row{{border-bottom:1px solid {C['outline']};transition:background .12s}}
.wpb-booked-table tbody tr.wpb-row:hover{{background:{C['surface_low']}}}
.wpb-booked-table tbody tr:last-child{{border-bottom:none}}
.wpb-booked-table td{{padding:12px 18px}}
.wpb-booked-empty{{padding:52px 20px;text-align:center;color:#9aa0a6;font-size:13px}}
.wpb-booked-empty .ico{{width:52px;height:52px;border-radius:50%;background:{C['surface_low']};
  display:flex;align-items:center;justify-content:center;font-size:22px;margin:0 auto 12px}}
.wpb-ministage-err{{background:#fbe4e4;color:{C['error']}}}
.wpb-ministage-pend{{background:#eef0f3;color:{C['on_var']}}}

/* ── Form fields: scoped + higher specificity so this always renders
     correctly regardless of any other .sp-fld rules on the page ── */
.wpb-form-fields{{margin-top:2px}}
.wpb-step-hdr{{display:flex;align-items:center;gap:10px;margin:24px 0 16px}}
.wpb-step-hdr:first-child{{margin-top:0}}
.wpb-step-num{{width:24px;height:24px;border-radius:50%;background:{C['dark']};color:#fff;font-size:11px;
  font-weight:800;display:flex;align-items:center;justify-content:center;flex-shrink:0}}
.wpb-step-title{{font-size:12.5px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:{C['dark']}}}
.wpb-form-fields .sp-row{{display:flex;gap:16px;flex-wrap:wrap}}
.wpb-form-fields .sp-row .sp-fld{{flex:1;min-width:200px}}
.wpb-form-fields .sp-fld{{margin-bottom:18px}}
.wpb-form-fields .sp-fld label{{display:block;font-size:11.5px;font-weight:600;color:{C['on_surface']};
  margin-bottom:7px;text-transform:none;letter-spacing:normal}}
.wpb-form-fields .sp-fld label.req:after{{content:" *";color:{C['error']};font-weight:700}}
.wpb-form-fields .sp-fld input[type=text],
.wpb-form-fields .sp-fld input[type=number],
.wpb-form-fields .sp-fld input[type=date],
.wpb-form-fields .sp-fld select,
.wpb-form-fields .sp-fld textarea{{
  box-sizing:border-box;width:100%;padding:11px 14px;border:1.5px solid {C['outline']};border-radius:9px;
  font-size:14px;background:#fff;color:{C['on_surface']};outline:none;font-family:inherit;
  transition:border-color .15s,box-shadow .15s}}
.wpb-form-fields .sp-fld select{{
  appearance:none;-webkit-appearance:none;-moz-appearance:none;cursor:pointer;padding-right:38px;
  background-image:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='10' height='6'><path d='M0 0l5 6 5-6z' fill='%2343474e'/></svg>");
  background-repeat:no-repeat;background-position:right 14px center}}
.wpb-form-fields .sp-fld input:focus,
.wpb-form-fields .sp-fld select:focus,
.wpb-form-fields .sp-fld textarea:focus{{
  border-color:{C['dark']};box-shadow:0 0 0 3px rgba(0,14,36,.08)}}
.wpb-form-fields .sp-fld textarea{{resize:vertical;min-height:80px;line-height:1.5}}
.wpb-amt-wrap{{position:relative}}
.wpb-amt-wrap span{{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:{C['on_var']};font-size:14px;pointer-events:none}}
.wpb-amt-wrap input{{padding-left:30px !important}}
.wpb-mode-switch{{display:inline-flex;border:1.5px solid {C['outline']};border-radius:9px}}
.wpb-mode-btn{{padding:9px 18px;font-size:13px;font-weight:700;border:none;background:#fff;
  color:{C['on_var']};cursor:pointer;transition:background .15s,color .15s}}
.wpb-mode-item:first-child .wpb-mode-btn{{border-radius:7.5px 0 0 7.5px;border-right:1.5px solid {C['outline']}}}
.wpb-mode-item:last-child .wpb-mode-btn{{border-radius:0 7.5px 7.5px 0}}
.wpb-mode-btn.active{{background:{C['dark']};color:#fff}}
</style>"""

    html_body = f"""
<div class="main"><div class="canvas" style="max-width:1100px">
  <div style="margin-bottom:14px">
    <h2 style="font-size:24px;font-weight:800;color:{C['dark']}">Weekly Payment Booking</h2>
    <p style="font-size:13px;color:#6b7280;margin-top:4px">Book vendor / sub-contractor payments for Procurement and Accounts verification.</p>
  </div>
  <div class="wpb-period-row">
    <div class="wpb-period">📅 Booking Period: {period_start} to {period_end}</div>
    <div class="wpb-mini-stat">Booked this period<b> ₹ {amt_period:,.0f}</b></div>
    <div class="wpb-mini-stat">Pending<b> {n_pending}</b></div>
    <div class="wpb-mini-stat">Approved<b> {n_approved}</b></div>
  </div>

  {alert_html}

  <div class="wpb-card">
   <div class="wpb-form-fields">
    <div class="wpb-step-hdr"><span class="wpb-step-num">1</span><span class="wpb-step-title">Project</span></div>
    <div class="sp-row" id="wpb-mode-row" style="display:{'none' if is_indirect else 'flex'};align-items:flex-start">
      <div class="sp-fld" id="wpb-mode-switch-wrap" style="flex:0 0 auto">
        <label>Book Against</label>
        <div class="wpb-mode-switch">
          <div class="wpb-mode-item">
            <button type="button" id="wpb-mode-project-btn" class="wpb-mode-btn active" onclick="wpbSetBookMode('project')">🏗 Project</button>
          </div>
          <div class="wpb-mode-item">
            <button type="button" id="wpb-mode-location-btn" class="wpb-mode-btn" onclick="wpbSetBookMode('location')">📍 Location</button>
          </div>
        </div>
      </div>
      <div class="sp-fld" id="wpb-proj-fld" style="flex:1;min-width:200px;display:block">
        <form method="get" action="/weekly_booking" id="wpb-proj-select-form">
          <input type="hidden" name="is_indirect" value="{'1' if is_indirect else '0'}">
          <label class="req">Select Project</label>
          <select name="project_id" id="wpb-project-select" onchange="this.form.submit()">{proj_opts}</select>
        </form>
      </div>
      <div class="sp-fld" id="wpb-location-fld" style="flex:1;min-width:200px;display:none">
        <label class="req">Location</label>
        <select name="location" id="wpb-location-select" form="wpb-form">
          <option value="">— Select location —</option>
          {"".join(f'<option value="{loc}" {"selected" if loc==ef.get("location") else ""}>{loc}</option>' for loc in BOOKING_LOCATIONS)}
        </select>
      </div>
    </div>
    <input type="hidden" name="book_mode" id="wpb-book-mode-hidden" value="project" form="wpb-form">
    {head_toggle}
    {info_card}
    {warn_html}

    <form method="post" action="/save_booking" enctype="multipart/form-data" id="wpb-form"
          {"onsubmit=\"return wpbGuardSubmit()\"" if guard_blocked else ""}>
      <input type="hidden" name="is_indirect" value="{'1' if is_indirect else '0'}">
      <input type="hidden" name="project_id" id="wpb-project-id-hidden" value="{sel_project_id}">
      <input type="hidden" name="eac_bypass" id="wpb-eac-bypass" value="0">
      {f'<input type="hidden" name="booking_id" value="{edit_id_val}">' if edit_booking else ""}

      <div class="wpb-step-hdr"><span class="wpb-step-num">2</span><span class="wpb-step-title">Payment Details{' — Editing '+html.escape(str(ef.get('booking_uid','')))+' (was disapproved)' if edit_booking else ''}</span></div>
      <div class="sp-row">
        <div class="sp-fld"><label class="req">Payee Name</label><input type="text" name="payee_name" value="{_ev('payee_name')}" placeholder="e.g. ABC Traders" required></div>
        <div class="sp-fld"><label class="req">Payment Category</label>
          <select name="payment_category" required><option value="">— Select —</option>{"".join(f'<option value="{c}" {"selected" if c==ef.get("payment_category") else ""}>{c}</option>' for c in PAYMENT_CATEGORIES)}</select>
        </div>
      </div>
      <div class="sp-row">
        <div class="sp-fld"><label class="req">Requested Amount</label>
          <div class="wpb-amt-wrap"><span>₹</span><input type="number" step="0.01" name="requested_amount" value="{_ev('requested_amount') if edit_booking else ''}" placeholder="0.00" required></div>
        </div>
        <div class="sp-fld"><label class="req">Invoice Date</label><input type="date" name="invoice_date" value="{_ev('invoice_date')}" required></div>
      </div>
      <div class="sp-fld"><label class="req">Recommendation / Necessity</label>
        <textarea name="justification" rows="3" placeholder="Why is this payment necessary?" required>{_ev('justification')}</textarea>
      </div>

      <div class="wpb-step-hdr"><span class="wpb-step-num">3</span><span class="wpb-step-title">Necessary Documents</span></div>
      <div class="wpb-docs">{doc_inputs}</div>
      {'<p style="font-size:11.5px;color:#6b7280;margin-top:6px">Previously attached documents stay on file — attach here only if you need to add or replace any.</p>' if edit_booking else ''}

      <div style="margin-top:22px;display:flex;gap:10px">
        <button type="submit" id="wpb-submit-btn" class="btn-primary" style="padding:11px 24px;font-size:13.5px;border-radius:9px" {"disabled" if guard_blocked else ""}>{'Resend for Approval' if edit_booking else 'Submit for Approval'}</button>
      </div>
    </form>
   </div>
  </div>

  <div class="wpb-booked-card">
    <div class="wpb-booked-hdr"><h3>List of Payments Booked</h3><span class="wpb-booked-count">{len(my_bookings)} entries</span></div>
    <div class="wpb-booked-table-wrap">
      <table class="wpb-booked-table">
        <thead><tr><th>Booking ID</th><th>Project</th><th>Location</th><th>Payee</th><th class="text-right">Amount</th><th>Date</th><th>Status</th><th>Actions</th></tr></thead>
        <tbody>{rows_html}</tbody>
      </table>
    </div>
  </div>
</div></div>

<script>
function toggleNotes(id){{
  var row = document.getElementById('notes-'+id);
  row.style.display = (row.style.display === 'none') ? 'table-row' : 'none';
}}
function wpbSetBookMode(mode){{
  var isProj = mode === 'project';
  document.getElementById('wpb-proj-fld').style.display = isProj ? 'block' : 'none';
  document.getElementById('wpb-location-fld').style.display = isProj ? 'none' : 'block';
  document.getElementById('wpb-mode-project-btn').classList.toggle('active', isProj);
  document.getElementById('wpb-mode-location-btn').classList.toggle('active', !isProj);
  document.getElementById('wpb-book-mode-hidden').value = mode;
  var locSel = document.getElementById('wpb-location-select');
  var projSel = document.getElementById('wpb-project-select');
  if (isProj) {{
    locSel.required = false;
    locSel.value = '';
    projSel.disabled = false;
  }} else {{
    locSel.required = true;
    document.getElementById('wpb-project-id-hidden').value = '';
    if (projSel) projSel.value = '';
  }}
}}
function wpbGuardSubmit(){{
  // Blocks submission until the EAC bypass button has been clicked.
  return document.getElementById('wpb-eac-bypass') &&
         document.getElementById('wpb-eac-bypass').value === '1';
}}
function wpbBypass(){{
  document.getElementById('wpb-eac-bypass').value = '1';
  document.getElementById('wpb-submit-btn').disabled = false;
  var alertBox = document.getElementById('wpb-eac-alert');
  alertBox.innerHTML = '⚠ EAC check bypassed. This booking will be flagged for review.';
}}
function wpbShowFiles(input, targetId){{
  var el = document.getElementById(targetId);
  if (!el) return;
  if (input.files && input.files.length) {{
    el.textContent = input.files.length === 1 ? input.files[0].name : (input.files.length + ' files selected');
    el.style.color = '{C["secondary"]}';
    el.style.fontWeight = '700';
  }} else {{
    el.textContent = 'Click to attach';
    el.style.color = '';
    el.style.fontWeight = '';
  }}
}}
</script>
</body></html>"""

    return _html_head("Weekly Payment Booking") + css + "<body>" + \
           _payments_sidebar("weekly_booking") + _topbar("WEEKLY PAYMENT BOOKING", username, portal="payments") + html_body


def payments_dashboard(username="Admin", role="pm") -> str:
    """Landing page for the standalone Payments portal — quick stats plus
    shortcuts into Weekly Payment Booking and Payment Approvals."""
    period_start, period_end = get_booking_period()
    stats = fetch_booking_stats()
    my_bookings = fetch_bookings_for_pm(username) if is_booking_role(role) else []
    amt_period = sum(_n(b.get("requested_amount")) for b in my_bookings
                      if str(b.get("period_start", "")) == period_start)

    show_approvals = role not in ("pm", "head")

    cards = f"""
    <a href="/weekly_booking" class="pd-card">
      <div class="pd-card-ico">🗓️</div>
      <div class="pd-card-title">Weekly Payment Booking</div>
      <div class="pd-card-sub">Book vendor / sub-contractor payments for Procurement and Accounts verification.</div>
      <div class="pd-card-cta">Open →</div>
    </a>"""
    if show_approvals:
        cards += f"""
    <a href="/payment_approvals" class="pd-card">
      <div class="pd-card-ico">✅</div>
      <div class="pd-card-title">Payment Approvals</div>
      <div class="pd-card-sub">Review and act on pending bookings across Procurement, Accounts and Director stages.</div>
      <div class="pd-card-cta">Open →</div>
    </a>"""

    css = f"""<style>
.pd-stats{{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:26px}}
.pd-stat{{background:#fff;border:1px solid {C['outline']};border-radius:12px;padding:16px 22px;min-width:150px}}
.pd-stat-lbl{{font-size:11px;color:{C['on_var']};text-transform:uppercase;letter-spacing:.04em;margin-bottom:6px}}
.pd-stat-val{{font-size:22px;font-weight:800;color:{C['dark']}}}
.pd-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:18px}}
.pd-card{{background:#fff;border:1px solid {C['outline']};border-radius:14px;padding:26px;text-decoration:none;
  display:block;box-shadow:0 1px 3px rgba(0,0,0,.04);transition:box-shadow .15s,transform .15s}}
.pd-card:hover{{box-shadow:0 6px 18px rgba(0,14,36,.1);transform:translateY(-2px)}}
.pd-card-ico{{font-size:28px;margin-bottom:12px}}
.pd-card-title{{font-size:16px;font-weight:800;color:{C['dark']};margin-bottom:8px}}
.pd-card-sub{{font-size:12.5px;color:{C['on_var']};line-height:1.5;margin-bottom:16px}}
.pd-card-cta{{font-size:12px;font-weight:700;color:{C['secondary']}}}
</style>"""

    html_body = f"""
<div class="main"><div class="canvas" style="max-width:1100px">
  <div style="margin-bottom:14px">
    <h2 style="font-size:24px;font-weight:800;color:{C['dark']}">Payments Dashboard</h2>
    <p style="font-size:13px;color:#6b7280;margin-top:4px">Weekly payment bookings and approval status at a glance.</p>
  </div>

  <div class="pd-stats">
    <div class="pd-stat"><div class="pd-stat-lbl">Booking Period</div><div class="pd-stat-val" style="font-size:15px">{period_start} – {period_end}</div></div>
    <div class="pd-stat"><div class="pd-stat-lbl">Booked This Period</div><div class="pd-stat-val">₹ {amt_period:,.0f}</div></div>
    <div class="pd-stat"><div class="pd-stat-lbl">Pending</div><div class="pd-stat-val">{stats.get('pending', 0)}</div></div>
    <div class="pd-stat"><div class="pd-stat-lbl">Approved</div><div class="pd-stat-val">{stats.get('approved', 0)}</div></div>
  </div>

  <div class="pd-grid">{cards}</div>
</div></div>
</body></html>"""

    return _html_head("Payments Dashboard") + css + "<body>" + \
           _payments_sidebar("payments_dashboard") + _topbar("PAYMENTS DASHBOARD", username, portal="payments") + html_body


def payment_approvals(alert_html="", username="Admin", role="procurement", selected_id=None) -> str:
    bookings = fetch_all_bookings(include_archived=False)
    stats = fetch_booking_stats()
    sel = fetch_booking_detail(selected_id) if selected_id else {}

    STAGE_ICONS = {"pending": "⏳", "approved": "✓", "disapproved": "✕"}

    rows_html = ""
    for b in bookings:
        active_cls = "wpb-list-item active" if selected_id and int(selected_id) == b["id"] else "wpb-list-item"
        status = (b.get("overall_status") or "pending").lower()
        badge_cls = {"approved": "ok", "disapproved": "err"}.get(status, "pend")
        # which stage is currently awaiting action, for a quick-glance dot
        if (b.get("procurement_status") or "pending") == "pending":
            stage_lbl, stage_dot = "Awaiting Procurement", "🧾"
        elif (b.get("accounts_status") or "pending") == "pending":
            stage_lbl, stage_dot = "Awaiting Accounts", "💳"
        elif (b.get("director_status") or "pending") == "pending":
            stage_lbl, stage_dot = "Awaiting Director", "🖋️"
        else:
            stage_lbl, stage_dot = status.title(), ""
        # which stage disapproved it, for a quick-glance "who" on the list row
        disapprover = ""
        if status == "disapproved":
            for _by_key, _status_key in (
                ("procurement_by", "procurement_status"),
                ("accounts_by",    "accounts_status"),
                ("director_by",    "director_status"),
            ):
                if (b.get(_status_key) or "").lower() == "disapproved":
                    disapprover = b.get(_by_key) or ""
                    break
        rows_html += f"""
        <a href="/payment_approvals?id={b['id']}" class="{active_cls}" data-status="{status}">
          <div class="wpb-li-top">
            <span class="wpb-li-uid">{b.get('booking_uid','')}</span>
            <span class="wpb-badge wpb-badge-{badge_cls}">{status.title()}</span>
          </div>
          <div class="wpb-li-proj">{b.get('project_label','')}</div>
          <div class="wpb-li-meta">{b.get('pm_name','')} · ₹ {_n(b.get('requested_amount')):,.0f}</div>
          <div class="wpb-li-meta" style="margin-top:2px">📅 Invoice: {b.get('invoice_date','') or '—'}</div>
          {f'<div class="wpb-li-stage">{stage_dot} {stage_lbl}</div>' if status == 'pending' else ''}
          {f'<div class="wpb-li-stage" style="color:{C["error"]}">✕ Disapproved by {html.escape(str(disapprover))}</div>' if status == 'disapproved' and disapprover else ''}
        </a>"""
    if not rows_html:
        rows_html = ('<div style="padding:40px 20px;text-align:center;color:#9aa0a6">'
                     '<div style="font-size:26px;margin-bottom:6px">✅</div>No active bookings to review.</div>')

    def _stage_block(stage_name, stage_key, note_key, status_key, by_key, at_key, decide_path, can_act, icon, is_last=False):
        stage_status = (sel.get(status_key) or "pending")
        note_val = sel.get(note_key) or ""
        by_val = sel.get(by_key) or ""
        at_val = str(sel.get(at_key) or "")[:16]
        dot_cls = {"approved": "ok", "disapproved": "err"}.get(stage_status, "pend")
        form_html = ""
        if can_act and stage_status == "pending":
            form_html = f"""
            <form method="post" action="{decide_path}" style="margin-top:10px">
              <input type="hidden" name="booking_id" value="{sel.get('id','')}">
              <textarea name="note" rows="2" placeholder="Verification note…"
                style="width:100%;padding:8px 10px;border:1px solid {C['outline']};border-radius:6px;font-size:12px;margin-bottom:8px"></textarea>
              <div style="display:flex;gap:8px">
                <button type="submit" name="decision" value="approve" class="btn-primary" style="background:#1a7a34">✓ Approve</button>
                <button type="submit" name="decision" value="disapprove" class="btn-primary" style="background:{C['error']}">✕ Disapprove</button>
              </div>
            </form>"""
        elif stage_status == "pending":
            form_html = f'<div class="wpb-stage-wait">Waiting on this stage…</div>'
        else:
            form_html = f"""<div class="wpb-stage-done">
              <b>{stage_status.title()}</b>{f' by {by_val} on {at_val}' if by_val else ''}
              {f'<div class="wpb-stage-note">“{note_val}”</div>' if note_val else ''}
              </div>"""
        return f"""
        <div class="wpb-stage{'' if not is_last else ' wpb-stage-last'}">
          <div class="wpb-stage-rail">
            <span class="wpb-stage-dot wpb-stage-dot-{dot_cls}">{STAGE_ICONS.get(stage_status,'⏳')}</span>
            {'' if is_last else '<span class="wpb-stage-line"></span>'}
          </div>
          <div class="wpb-stage-body">
            <div class="wpb-stage-name">{icon} {stage_name}</div>
            {form_html}
          </div>
        </div>"""

    procurement_status = sel.get("procurement_status") or "pending"
    accounts_status    = sel.get("accounts_status") or "pending"
    director_status    = sel.get("director_status") or "pending"
    can_procurement = (is_procurement_role(role) or is_user_role(role)) and (procurement_status == "pending")
    can_accounts = (is_accounts_role(role) or is_user_role(role)) and (procurement_status == "approved") and (accounts_status == "pending")
    can_director = (role == "director") and (accounts_status == "approved") and (director_status == "pending")

    detail_html = ('<div style="padding:70px 20px;text-align:center;color:#9aa0a6">'
                    '<div style="font-size:30px;margin-bottom:8px">👈</div>Select a booking from the list to review it.</div>')
    if sel:
        docs = sel.get("docs") or []
        docs_html = "".join(
            f'<span class="wpb-doc-chip">📎 {d.get("doc_type","")}: {html.escape(d.get("file_name","") or "")}</span>'
            for d in docs
        ) or '<span style="font-size:12px;color:#9aa0a6">No documents attached</span>'
        export_link = ""
        if sel.get("overall_status") == "approved":
            export_link = f'<a href="/download_stamped_doc?id={sel["id"]}" class="btn-outline">⬇ Stamped PDF</a>'
        overall = (sel.get("overall_status") or "pending").lower()
        overall_cls = {"approved": "ok", "disapproved": "err"}.get(overall, "pend")

        # ── Who disapproved (whichever stage rejected it first) ───────────────
        disapproved_by_html = ""
        if overall == "disapproved":
            for _by_key, _at_key, _note_key, _stage_lbl in (
                ("procurement_by", "procurement_at", "procurement_note", "Procurement"),
                ("accounts_by",    "accounts_at",    "accounts_note",    "Accounts"),
                ("director_by",    "director_at",    "director_note",    "Director"),
            ):
                if (sel.get(_by_key.replace("_by", "_status")) or "").lower() == "disapproved":
                    _by  = sel.get(_by_key) or "—"
                    _at  = str(sel.get(_at_key) or "")[:16]
                    _note = sel.get(_note_key) or ""
                    disapproved_by_html = (
                        f'<div class="wpb-disapproved-by">'
                        f'✕ Disapproved by <b>{html.escape(str(_by))}</b> ({_stage_lbl})'
                        f'{f" on {_at}" if _at else ""}'
                        f'{f"<div class=\'wpb-stage-note\'>“{html.escape(str(_note))}”</div>" if _note else ""}'
                        f'</div>'
                    )
                    break

        detail_html = f"""
        <div class="wpb-detail-hdr">
          <div>
            <div class="wpb-detail-title">{sel.get('booking_uid','')}</div>
            <div class="wpb-detail-sub">{sel.get('project_label','')}</div>
          </div>
          <span class="wpb-badge wpb-badge-{overall_cls}" style="font-size:11px">{overall.title()}</span>
        </div>
        {disapproved_by_html}
        <div class="wpb-detail-facts">
          <div class="wpb-fact"><span>Booked By</span><b>{sel.get('pm_name','')}</b></div>
          <div class="wpb-fact"><span>Amount</span><b>₹ {_n(sel.get('requested_amount')):,.0f}</b></div>
          <div class="wpb-fact"><span>Category</span><b>{sel.get('payment_category','')}</b></div>
          <div class="wpb-fact"><span>Invoice Date</span><b>{sel.get('invoice_date','')}</b></div>
        </div>
        {f'<div style="margin:12px 0">{export_link}</div>' if export_link else ''}
        <div class="wpb-detail-just">{sel.get('justification','')}</div>
        <div class="wpb-docs-row">{docs_html}</div>

        <div class="sp-section-lbl">Approval Workflow</div>
        <div class="wpb-stages">
          {_stage_block("Procurement Department", "procurement", "procurement_note", "procurement_status", "procurement_by", "procurement_at", "/procurement_decide", can_procurement, "🧾")}
          {_stage_block("Accounts Department", "accounts", "accounts_note", "accounts_status", "accounts_by", "accounts_at", "/accounts_decide", can_accounts, "💳")}
          {_stage_block("Director (Finance)", "director", "director_note", "director_status", "director_by", "director_at", "/director_decide", can_director, "🖋️", is_last=True)}
        </div>
        """

    css = f"""<style>
.wpb-filter-row{{display:flex;gap:8px;margin-bottom:12px}}
.wpb-filter-btn{{background:#fff;border:1px solid {C['outline']};border-radius:20px;padding:6px 14px;font-size:11.5px;font-weight:600;color:{C['on_var']};cursor:pointer}}
.wpb-filter-btn.active{{background:{C['dark']};color:#fff;border-color:{C['dark']}}}
.wpb-list-item{{display:block;padding:13px 16px;border-bottom:1px solid {C['outline']};text-decoration:none;color:inherit;transition:background .12s}}
.wpb-list-item:hover{{background:{C['surface_low']}}}
.wpb-list-item.active{{background:#f0f3fa;box-shadow:inset 3px 0 0 {C['dark']}}}
.wpb-li-top{{display:flex;justify-content:space-between;align-items:center;margin-bottom:3px}}
.wpb-li-uid{{font-weight:700;color:{C['dark']};font-size:13px}}
.wpb-li-proj{{font-size:12px;color:{C['on_surface']}}}
.wpb-li-meta{{font-size:11px;color:{C['on_var']};margin-top:2px}}
.wpb-li-stage{{font-size:10.5px;color:{C['secondary']};font-weight:700;margin-top:6px}}
.wpb-approvals-grid{{display:grid;grid-template-columns:320px 1fr;gap:20px;align-items:start}}
.wpb-list-card{{background:#fff;border:1px solid {C['outline']};border-radius:14px;overflow:hidden;max-height:74vh;overflow-y:auto}}
.wpb-detail-card{{background:#fff;border:1px solid {C['outline']};border-radius:14px;padding:24px;min-height:74vh}}
.wpb-detail-hdr{{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:14px}}
.wpb-detail-title{{font-weight:800;font-size:17px;color:{C['dark']}}}
.wpb-detail-sub{{font-size:12.5px;color:{C['on_var']};margin-top:2px}}
.wpb-detail-facts{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:12px}}
.wpb-fact{{background:{C['surface_low']};border-radius:8px;padding:9px 12px}}
.wpb-fact span{{display:block;font-size:9.5px;text-transform:uppercase;letter-spacing:.06em;color:{C['on_var']};margin-bottom:3px}}
.wpb-fact b{{font-size:13px;color:{C['dark']}}}
.wpb-detail-just{{font-size:12.5px;color:{C['on_surface']};background:#fbfbfc;border:1px solid {C['outline']};border-radius:8px;padding:10px 12px;margin-bottom:12px}}
.wpb-docs-row{{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:6px}}
.wpb-doc-chip{{background:{C['surface_low']};border:1px solid {C['outline']};border-radius:20px;padding:5px 12px;font-size:11px;color:{C['on_var']}}}
.wpb-stages{{margin-top:6px}}
.wpb-stage{{display:flex;gap:12px}}
.wpb-stage-rail{{display:flex;flex-direction:column;align-items:center}}
.wpb-stage-dot{{width:30px;height:30px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:13px;flex-shrink:0}}
.wpb-stage-dot-ok{{background:#1a7a34;color:#fff}}
.wpb-stage-dot-err{{background:{C['error']};color:#fff}}
.wpb-stage-dot-pend{{background:{C['surface_high']};color:{C['on_var']}}}
.wpb-stage-line{{width:2px;flex:1;background:{C['outline']};margin:2px 0}}
.wpb-stage-body{{flex:1;padding-bottom:18px}}
.wpb-stage-last .wpb-stage-body{{padding-bottom:0}}
.wpb-stage-name{{font-weight:700;color:{C['dark']};font-size:13px;padding-top:4px;margin-bottom:6px}}
.wpb-stage-wait{{font-size:11.5px;color:{C['on_var']};font-style:italic}}
.wpb-stage-done{{font-size:12px;color:{C['on_var']}}}
.wpb-stage-note{{margin-top:4px;font-style:italic;color:{C['on_surface']}}}
.wpb-stats-bar{{display:grid;grid-template-columns:repeat(3,1fr) auto;gap:14px;margin-bottom:18px;align-items:stretch}}
.wpb-stat-card{{background:#fff;border:1px solid {C['outline']};border-radius:12px;padding:14px 18px}}
.wpb-stat-card .stat-lbl{{font-size:10px;color:{C['on_var']};text-transform:uppercase;letter-spacing:.06em}}
.wpb-stat-card .stat-val{{font-size:22px;font-weight:800;color:{C['dark']};margin-top:2px}}
.wpb-stat-card.dark{{background:{C['dark']};border-color:{C['dark']}}}
.wpb-stat-card.dark .stat-lbl{{color:rgba(255,255,255,.65)}}
.wpb-stat-card.dark .stat-val{{color:#fff}}
.wpb-export-card{{display:flex;align-items:center;justify-content:center}}
.wpb-badge{{padding:3px 10px;border-radius:20px;font-size:10px;font-weight:700;text-transform:uppercase}}
.wpb-badge-ok{{background:#e3f5e6;color:#1a7a34}}
.wpb-badge-err{{background:#fbe4e4;color:{C['error']}}}
.wpb-badge-pend{{background:#fdf3da;color:{C['secondary']}}}
.wpb-disapproved-by{{background:#fbe4e4;color:{C['error']};border:1px solid #f3c8c8;border-radius:8px;padding:8px 12px;font-size:12px;margin-bottom:12px}}
.wpb-disapproved-by .wpb-stage-note{{margin-top:4px;font-style:italic;color:{C['error']};opacity:.85}}
</style>"""

    html_body = f"""
<div class="main"><div class="canvas" style="max-width:1200px">
  <div style="margin-bottom:14px">
    <h2 style="font-size:24px;font-weight:800;color:{C['dark']}">Payment Approvals</h2>
    <p style="font-size:13px;color:#6b7280;margin-top:4px">Procurement → Accounts → Director sign-off workflow.</p>
  </div>
  <div class="wpb-stats-bar">
    <div class="wpb-stat-card dark"><div class="stat-lbl">Total Bookings</div><div class="stat-val">{stats['total']}</div></div>
    <div class="wpb-stat-card"><div class="stat-lbl">Approved</div><div class="stat-val">{stats['approved']}</div></div>
    <div class="wpb-stat-card"><div class="stat-lbl">Pending</div><div class="stat-val">{stats['pending']}</div></div>
    <div class="wpb-stat-card wpb-export-card"><a href="/export_excel/payment_bookings" class="btn-primary">⬇ Export Excel</a></div>
  </div>
  {alert_html}
  <div class="wpb-filter-row">
    <span class="wpb-filter-btn active" onclick="wpbFilter(this,'all')">All</span>
    <span class="wpb-filter-btn" onclick="wpbFilter(this,'pending')">Pending</span>
    <span class="wpb-filter-btn" onclick="wpbFilter(this,'disapproved')">Disapproved</span>
  </div>
  <div class="wpb-approvals-grid">
    <div class="wpb-list-card" id="wpb-list-card">{rows_html}</div>
    <div class="wpb-detail-card">{detail_html}</div>
  </div>
</div></div>
<script>
function wpbFilter(btn, status){{
  document.querySelectorAll('.wpb-filter-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  document.querySelectorAll('#wpb-list-card > a').forEach(a=>{{
    var s = a.getAttribute('data-status');
    a.style.display = (status==='all' || s===status) ? '' : 'none';
  }});
}}
</script>
</body></html>"""

    return _html_head("Payment Approvals") + css + "<body>" + \
           _payments_sidebar("payment_approvals") + _topbar("PAYMENT APPROVALS", username, portal="payments") + html_body


# ═══════════════════════════════════════════════════════════════
#  HTTP HANDLER
# ═══════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════
#  MISSING EXPORT FUNCTIONS
# ═══════════════════════════════════════════════════════════════

def export_expenditures_excel_bytes() -> bytes:
    """Export all expenditure records as a formatted .xlsx file."""
    from openpyxl.styles import Font, PatternFill, Alignment
    projects = fetch_all_projects()

    wb = Workbook()
    ws = wb.active
    ws.title = "Expenditures"

    hdr_fill = PatternFill("solid", fgColor="000E24")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    alt_fill = PatternFill("solid", fgColor="F2F4F6")
    mono     = Font(name="Courier New", size=10)

    headers = [
        "Project ID", "CO.No", "Project Name", "LOB",
        "Data As Of", "ERP Update Date",
        "ITD COGS", "ITD DJCS", "ITD Equipment", "ITD Manpower", "ITD Vehicles", "ITD Others",
        "ERP COGS", "ERP DJCS", "ERP Equipment", "ERP Manpower", "ERP Vehicles", "ERP Others",
        "Total ITD", "Total ERP",
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    row_num = 2
    for p in projects:
        pid = p.get("id")
        exp = fetch_expenditure_for_co_no(p.get("co_no"))
        if not exp:
            continue

        itd_vals = [
            _n(exp.get("itd_cogs")),  _n(exp.get("itd_djcs")),
            _n(exp.get("itd_equip")), _n(exp.get("itd_manpower")),
            _n(exp.get("itd_vehicles")), _n(exp.get("itd_others")),
        ]
        erp_vals = [
            _n(exp.get("erp_cogs")),  _n(exp.get("erp_djcs")),
            _n(exp.get("erp_equip")), _n(exp.get("erp_manpower")),
            _n(exp.get("erp_vehicles")), _n(exp.get("erp_others")),
        ]
        total_itd = sum(itd_vals)
        total_erp = sum(erp_vals)

        row_data = [
            pid,
            p.get("co_no", ""),
            p.get("project_name", ""),
            p.get("hier_business_line") or p.get("business_line", ""),
            exp.get("data_as_of", ""),
            exp.get("erp_update_date", ""),
            *itd_vals,
            *erp_vals,
            total_itd,
            total_erp,
        ]
        ws.append(row_data)

        fill = alt_fill if row_num % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        for col_idx, cell in enumerate(ws[row_num], start=1):
            cell.fill = fill
            if col_idx >= 7:
                cell.font             = mono
                cell.number_format    = '#,##0.00'
                cell.alignment        = Alignment(horizontal="right")

        row_num += 1

    # Totals footer
    if row_num > 2:
        totals_label = ["TOTAL", "", "", "", "", ""]
        col_totals   = []
        for col_idx in range(7, 21):
            col_letter = ws.cell(1, col_idx).column_letter
            col_totals.append(
                f"=SUM({col_letter}2:{col_letter}{row_num - 1})"
            )
        ws.append(totals_label + col_totals)
        for cell in ws[row_num]:
            cell.font      = Font(bold=True, name="Courier New", size=10)
            cell.fill      = PatternFill("solid", fgColor="E6E8EA")
            cell.alignment = Alignment(horizontal="right")

    # Column widths
    col_widths = [8, 12, 38, 20, 14, 14,
                  14, 14, 14, 14, 14, 14,
                  14, 14, 14, 14, 14, 14,
                  14, 14]
    for idx, w in enumerate(col_widths, start=1):
        ws.column_dimensions[ws.cell(1, idx).column_letter].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

def _build_table_pdf(title: str, subtitle: str, headers: list,
                     rows: list, footer_label: str = "", footer_value: float = 0.0) -> bytes:
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    import io

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                            leftMargin=28, rightMargin=28,
                            topMargin=28, bottomMargin=28)
    styles = getSampleStyleSheet()
    dark   = colors.HexColor("#000e24")
    gold   = colors.HexColor("#775a19")
    light  = colors.HexColor("#f2f4f6")

    title_style = ParagraphStyle("t", parent=styles["Heading1"],
                                 fontSize=16, textColor=dark, spaceAfter=4)
    sub_style   = ParagraphStyle("s", parent=styles["Normal"],
                                 fontSize=9, textColor=colors.HexColor("#43474e"), spaceAfter=10)

    table_data = [headers] + (rows or [["No data"]])
    if footer_label:
        table_data.append([footer_label] + [""] * (len(headers) - 2) +
                          [f"{footer_value:,.0f}"])

    col_count = len(headers)
    page_w    = landscape(A4)[0] - 56
    col_w     = page_w / col_count

    tbl = Table(table_data, colWidths=[col_w] * col_count, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, 0),  dark),
        ("TEXTCOLOR",    (0, 0), (-1, 0),  colors.white),
        ("FONTSIZE",     (0, 0), (-1, 0),  8),
        ("FONTNAME",     (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("ROWBACKGROUNDS",(0,1), (-1,-1),   [colors.white, light]),
        ("FONTSIZE",     (0, 1), (-1, -1), 8),
        ("TEXTCOLOR",    (0, 1), (-1, -1), dark),
        ("GRID",         (0, 0), (-1, -1), 0.3, colors.HexColor("#c4c6d0")),
        ("TOPPADDING",   (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 5),
        ("LEFTPADDING",  (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("ALIGN",        (0, 0), (-1, -1), "LEFT"),
        ("BACKGROUND",   (0, -1), (-1, -1), light) if footer_label else ("BACKGROUND", (0,0),(0,0), dark),
        ("FONTNAME",     (0, -1), (-1, -1), "Helvetica-Bold") if footer_label else ("FONTNAME",(0,0),(0,0),"Helvetica-Bold"),
    ]))

    story = [
        Paragraph(title, title_style),
        Paragraph(subtitle, sub_style),
        tbl,
    ]
    doc.build(story)
    buf.seek(0)
    return buf.read()


def export_excel_bytes() -> bytes:
    """Export all projects as a formatted .xlsx file."""
    projects = fetch_all_projects()
    wb = Workbook()
    ws = wb.active
    ws.title = "Projects"

    from openpyxl.styles import Font, PatternFill, Alignment
    hdr_fill = PatternFill("solid", fgColor="000E24")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    alt_fill = PatternFill("solid", fgColor="F2F4F6")

    headers = [
        "ID", "CO.No", "Project Name", "Business Line", "Project Manager",
        "Client Name", "Order Number", "Order Issue Date", "Order Value",
        "Variation Amt", "Revised OV", "Project Status", "Physical %",
        "Billed Amount", "Collection", "ITD Cost", "Baseline EAC",
        "Rev Latest EAC", "Project Type", "Parent Project",
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for i, p in enumerate(projects, start=2):
        ws.append([
            p.get("id"), p.get("co_no"), p.get("project_name"),
            p.get("business_line"), p.get("project_manager"), p.get("client_name"),
            p.get("order_number"), p.get("order_issue_date"),
            _n(p.get("order_value")), _n(p.get("variation_amt")),
            _n(p.get("revised_order_value")), p.get("project_status"),
            _n(p.get("physical_completion")), _n(p.get("billed_amount")),
            _n(p.get("collection_till_date")), _n(p.get("itd_cost")),
            _n(p.get("baseline_eac")), _n(p.get("rev_latest_eac")),
            p.get("project_type"), p.get("parent_project"),
        ])
        if i % 2 == 0:
            for cell in ws[i]:
                cell.fill = alt_fill

    col_widths = [6, 12, 40, 18, 20, 24, 16, 14, 14, 14, 14,
                  14, 10, 14, 14, 14, 14, 14, 10, 24]
    for idx, w in enumerate(col_widths, start=1):
        ws.column_dimensions[ws.cell(1, idx).column_letter].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def export_full_portfolio_excel_bytes() -> bytes:
    """Export ONE complete, comprehensive workbook covering every field and every
    entity used across the whole portal: Projects (all DB fields + computed metrics),
    Expenditures (ITD & ERP heads + variance), ERP Updates (detailed line items),
    Expense Entries, Billing & Collection (ledger + status), Variations,
    Payment Bookings (all approval stages, notes, and details), Payment Booking Docs,
    Payment Approval Logs, Documents & DPR, and all Master Reference Data tables.
    Each as its own sheet with professionally styled headers, freeze panes,
    and automatic column sizing."""
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    wb.remove(wb.active)

    hdr_fill = PatternFill("solid", fgColor="000E24")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    alt_fill = PatternFill("solid", fgColor="F8FAFC")

    def add_custom_sheet(title: str, rows: list, default_headers: list = None) -> None:
        ws = wb.create_sheet(title=title[:31])
        rows = [dict(r) for r in rows] if rows else []

        if rows:
            headers, seen = [], set()
            if default_headers:
                for h in default_headers:
                    headers.append(h)
                    seen.add(h)
            for r in rows:
                for k in r.keys():
                    if k not in seen:
                        headers.append(k)
                        seen.add(k)
        else:
            headers = default_headers or ["No Data"]

        special_labels = {
            "id": "ID", "co_no": "CO.No", "lob": "LOB", "ra_no": "RA No",
            "itd_cost": "ITD Cost (₹)", "ebf_1_35x": "EBF 1.35x (₹)",
            "itd_cogs": "ITD COGS (₹)", "itd_djcs": "ITD DJCS (₹)",
            "itd_equip": "ITD Equipment (₹)", "itd_manpower": "ITD Manpower (₹)",
            "itd_vehicles": "ITD Vehicles (₹)", "itd_others": "ITD Others (₹)",
            "erp_cogs": "ERP COGS (₹)", "erp_djcs": "ERP DJCS (₹)",
            "erp_equip": "ERP Equipment (₹)", "erp_manpower": "ERP Manpower (₹)",
            "erp_vehicles": "ERP Vehicles (₹)", "erp_others": "ERP Others (₹)",
            "total_itd": "Total ITD (₹)", "total_erp": "Total ERP (₹)",
            "variance_erp_vs_itd": "Variance ERP vs ITD (₹)",
            "billed_amount": "Billed Amount (₹)", "billed_amt": "Billed Amount (₹)",
            "collection_till_date": "Collection Till Date (₹)", "collection_amt": "Collection Amount (₹)",
            "pending_amt": "Pending Amount (₹)", "order_value": "Order Value (₹)",
            "variation_amt": "Variation Amount (₹)", "revised_order_value": "Revised Order Value (₹)",
            "computed_revised_order_value": "Computed Revised Order Value (₹)",
            "unbilled_amount": "Unbilled Amount (₹)", "uncollected_amount": "Uncollected Amount (₹)",
            "balance_order_value": "Balance Order Value (₹)", "estimate_to_complete_etc": "Estimate to Complete ETC (₹)",
            "baseline_eac": "Baseline EAC (₹)", "rev_latest_eac": "Rev Latest EAC (₹)",
            "eac_variance": "EAC Variance (₹)", "booking_uid": "Booking UID",
            "pm_name": "PM Name", "pm_username": "PM Username",
        }

        def format_header(h):
            if h in special_labels:
                return special_labels[h]
            return h.replace("_", " ").title()

        ws.append([format_header(h) for h in headers])
        for cell in ws[1]:
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 26

        if rows:
            for i, r in enumerate(rows, start=2):
                row_vals = []
                for h in headers:
                    v = r.get(h)
                    if isinstance(v, float):
                        v = round(v, 2)
                    elif isinstance(v, (datetime, date)):
                        v = str(v)
                    row_vals.append(v if v is not None else "")
                ws.append(row_vals)
                if i % 2 == 0:
                    for cell in ws[i]:
                        cell.fill = alt_fill
        else:
            ws.append(["(No records in database)" if h == headers[0] else "" for h in headers])

        for idx, h in enumerate(headers, start=1):
            hdr_len = len(format_header(h))
            ws.column_dimensions[get_column_letter(idx)].width = max(13, min(42, hdr_len + 4))
        ws.freeze_panes = "A2"

    try:
        with ENGINE.connect() as conn:
            # 1. Projects (Enriched with all DB fields and all calculated metrics)
            proj_rows = _rows_to_dicts(conn, "SELECT * FROM projects ORDER BY id DESC")
            enriched_projects = []
            for p in proj_rows:
                rev_ov = _rev_ov(p)
                billed = _n(p.get("billed_amount"))
                coll = _n(p.get("collection_till_date"))
                itd = _n(p.get("itd_cost"))
                base_eac = _n(p.get("baseline_eac"))
                rev_eac = _n(p.get("rev_latest_eac"))
                eac = rev_eac if rev_eac > 0 else base_eac

                unbilled = round(max(rev_ov - billed, 0), 2)
                uncollected = round(max(billed - coll, 0), 2)
                balance_ov = round(max(rev_ov - coll, 0), 2)
                exp_ov_pct = round((itd / rev_ov * 100), 2) if rev_ov > 0 else 0.0
                bill_pct = round((billed / rev_ov * 100), 2) if rev_ov > 0 else 0.0
                coll_pct = round((coll / billed * 100), 2) if billed > 0 else 0.0
                ebf = round(itd * 1.35, 2)

                sched = _sched_row_metrics(p)
                eac_m = _eac_row_metrics(p)
                eac_lbl = EAC_STATUS_BADGES.get(eac_m.get("status_key"), ("", "", "On Track"))[2]

                row = dict(p)
                row["computed_revised_order_value"] = rev_ov
                row["unbilled_amount"] = unbilled
                row["billing_percentage"] = bill_pct
                row["uncollected_amount"] = uncollected
                row["collection_percentage"] = coll_pct
                row["balance_order_value"] = balance_ov
                row["ebf_1_35x"] = ebf
                row["expenditure_over_order_value_pct"] = exp_ov_pct
                row["estimate_to_complete_etc"] = eac_m.get("etc", 0.0)
                row["eac_variance"] = round(eac - itd, 2) if eac > 0 else 0.0
                row["eac_status"] = eac_lbl
                row["wo_period_months"] = sched.get("period_months")
                row["project_end_date_computed"] = str(sched.get("proj_end_dt") or "")[:10]
                row["pct_time_elapsed"] = sched.get("pct_elapsed")
                row["pct_time_overrun"] = sched.get("pct_overrun")
                row["schedule_status"] = "Overrun" if sched.get("is_overrun") else ("Completed" if (sched.get("phy_pct") or 0) >= 100 else "On Track")
                enriched_projects.append(row)

            add_custom_sheet("Projects", enriched_projects)

            # 2. Expenditures
            exp_sql = """
                SELECT e.id, e.project_id, e.co_no, p.project_name, e.lob, p.project_manager, p.client_name,
                       e.data_as_of, e.erp_update_date,
                       e.itd_cogs, e.itd_djcs, e.itd_equip, e.itd_manpower, e.itd_vehicles, e.itd_others,
                       (COALESCE(e.itd_cogs,0) + COALESCE(e.itd_djcs,0) + COALESCE(e.itd_equip,0) + 
                        COALESCE(e.itd_manpower,0) + COALESCE(e.itd_vehicles,0) + COALESCE(e.itd_others,0)) AS total_itd,
                       e.erp_cogs, e.erp_djcs, e.erp_equip, e.erp_manpower, e.erp_vehicles, e.erp_others,
                       (COALESCE(e.erp_cogs,0) + COALESCE(e.erp_djcs,0) + COALESCE(e.erp_equip,0) + 
                        COALESCE(e.erp_manpower,0) + COALESCE(e.erp_vehicles,0) + COALESCE(e.erp_others,0)) AS total_erp,
                       ((COALESCE(e.erp_cogs,0) + COALESCE(e.erp_djcs,0) + COALESCE(e.erp_equip,0) + 
                         COALESCE(e.erp_manpower,0) + COALESCE(e.erp_vehicles,0) + COALESCE(e.erp_others,0)) -
                        (COALESCE(e.itd_cogs,0) + COALESCE(e.itd_djcs,0) + COALESCE(e.itd_equip,0) + 
                         COALESCE(e.itd_manpower,0) + COALESCE(e.itd_vehicles,0) + COALESCE(e.itd_others,0))) AS variance_erp_vs_itd,
                       e.created_at
                FROM expenditures e
                LEFT JOIN projects p ON p.id = e.project_id OR (e.co_no IS NOT NULL AND e.co_no != '' AND p.co_no = e.co_no)
                ORDER BY e.id DESC
            """
            add_custom_sheet("Expenditures", _rows_to_dicts(conn, exp_sql))

            # 3. ERP Updates (All detailed line items)
            erp_sql = """
                SELECT eu.id, eu.project_id, eu.co_no, p.project_name, p.hier_business_line AS lob,
                       p.project_manager, eu.category, eu.amount, eu.created_by, eu.created_at
                FROM erp_updates eu
                LEFT JOIN projects p ON p.id = eu.project_id OR (eu.co_no IS NOT NULL AND eu.co_no != '' AND p.co_no = eu.co_no)
                ORDER BY eu.id DESC
            """
            add_custom_sheet("ERP Updates", _rows_to_dicts(conn, erp_sql))

            # 4. Expense Entries (Line items)
            ee_sql = """
                SELECT ee.id, ee.expenditure_id, ee.project_id, ee.co_no, p.project_name,
                       ee.category, ee.description, ee.quantity, ee.unit_cost, ee.line_total, ee.created_at
                FROM expense_entries ee
                LEFT JOIN projects p ON p.id = ee.project_id OR (ee.co_no IS NOT NULL AND ee.co_no != '' AND p.co_no = ee.co_no)
                ORDER BY ee.id DESC
            """
            add_custom_sheet("Expense Entries", _rows_to_dicts(conn, ee_sql),
                             default_headers=["id", "expenditure_id", "project_id", "co_no", "project_name",
                                              "category", "description", "quantity", "unit_cost", "line_total", "created_at"])

            # 5. Billing & Collection Ledger
            bc_sql = """
                SELECT bc.id, bc.project_id, p.co_no, p.project_name, p.hier_business_line AS lob,
                       p.project_manager, p.client_name, bc.ra_no, bc.billing_date, bc.billed_amt,
                       bc.collection_date, bc.collection_amt,
                       (COALESCE(bc.billed_amt,0) - COALESCE(bc.collection_amt,0)) AS pending_amt,
                       CASE 
                           WHEN COALESCE(bc.collection_amt,0) >= COALESCE(bc.billed_amt,0) AND COALESCE(bc.billed_amt,0) > 0 THEN 'Fully Collected'
                           WHEN COALESCE(bc.collection_amt,0) > 0 THEN 'Partially Collected'
                           ELSE 'Pending Collection'
                       END AS collection_status,
                       bc.created_at
                FROM billing_collection bc
                LEFT JOIN projects p ON p.id = bc.project_id
                ORDER BY bc.id DESC
            """
            add_custom_sheet("Billing & Collection", _rows_to_dicts(conn, bc_sql))

            # 6. Variations
            var_sql = "SELECT * FROM variations ORDER BY id DESC"
            add_custom_sheet("Variations", _rows_to_dicts(conn, var_sql),
                             default_headers=["id", "project_id", "child_co_no", "child_name", "parent_name",
                                              "original_value", "variation_amt", "revised_value", "variation_date", "status", "created_at"])

            # 7. Payment Bookings (Workflow & Approvals)
            pb_sql = """
                SELECT pb.id, pb.booking_uid, pb.project_label, pb.project_id, p.co_no, p.project_name,
                       pb.project_manager, pb.pm_username, pb.pm_name, pb.payee_name, pb.payment_category,
                       pb.requested_amount, pb.invoice_date, pb.justification, pb.location,
                       pb.period_start, pb.period_end, pb.is_indirect, pb.indirect_dept,
                       pb.overall_status, pb.status,
                       pb.procurement_status, pb.procurement_note, pb.procurement_by, pb.procurement_at,
                       pb.accounts_status, pb.accounts_note, pb.accounts_by, pb.accounts_at,
                       pb.director_status, pb.director_note, pb.director_by, pb.director_at,
                       pb.digitally_signed_by, pb.disapproval_reason, pb.archived, pb.created_by, pb.created_at, pb.updated_at
                FROM payment_bookings pb
                LEFT JOIN projects p ON p.id = pb.project_id
                ORDER BY pb.id DESC
            """
            add_custom_sheet("Payment Bookings", _rows_to_dicts(conn, pb_sql))

            # 8. Payment Booking Docs
            pbd_sql = """
                SELECT pbd.id, COALESCE(pbd.booking_id, pbd.payment_booking_id) AS booking_id,
                       pb.booking_uid, pb.payee_name,
                       COALESCE(pbd.file_name, pbd.doc_name) AS file_name,
                       COALESCE(pbd.file_path, pbd.doc_path) AS file_path,
                       pbd.doc_type, pbd.is_digitally_signed, pbd.approval_stamp,
                       COALESCE(pbd.uploaded_at, pbd.created_at) AS uploaded_at
                FROM payment_booking_docs pbd
                LEFT JOIN payment_bookings pb ON pb.id = COALESCE(pbd.booking_id, pbd.payment_booking_id)
                ORDER BY pbd.id DESC
            """
            add_custom_sheet("Payment Booking Docs", _rows_to_dicts(conn, pbd_sql))

            # 9. Payment Approval Logs
            pal_sql = """
                SELECT pal.id, pal.payment_booking_id, pb.booking_uid, pal.stage, pal.action,
                       pal.notes, pal.verified_by, pal.verified_at
                FROM payment_approval_logs pal
                LEFT JOIN payment_bookings pb ON pb.id = pal.payment_booking_id
                ORDER BY pal.id DESC
            """
            add_custom_sheet("Payment Approval Logs", _rows_to_dicts(conn, pal_sql),
                             default_headers=["id", "payment_booking_id", "booking_uid", "stage", "action", "notes", "verified_by", "verified_at"])

            # 10. Documents & DPR Documents
            docs_sql = """
                SELECT d.id, d.project_id, p.co_no, p.project_name, d.doc_name, d.doc_category, d.doc_link, d.created_at
                FROM documents d
                LEFT JOIN projects p ON p.id = d.project_id
                ORDER BY d.id DESC
            """
            add_custom_sheet("Documents", _rows_to_dicts(conn, docs_sql),
                             default_headers=["id", "project_id", "co_no", "project_name", "doc_name", "doc_category", "doc_link", "created_at"])

            try:
                dpr_sql = "SELECT hash, filename, created_at, updated_at FROM dpr_documents ORDER BY updated_at DESC"
                add_custom_sheet("DPR Documents", _rows_to_dicts(conn, dpr_sql),
                                 default_headers=["hash", "filename", "created_at", "updated_at"])
            except Exception as e_dpr:
                print(f"[export_full_portfolio_excel_bytes] dpr_documents: {e_dpr}")

            # 11. Master & Reference Data
            add_custom_sheet("Ref - Lines of Business", _rows_to_dicts(conn, "SELECT id, name, active FROM ref_lob ORDER BY id"))
            add_custom_sheet("Ref - Clients", _rows_to_dicts(conn, "SELECT id, name, active FROM ref_clients ORDER BY id"))
            add_custom_sheet("Ref - Project Managers", _rows_to_dicts(conn, "SELECT id, name, email, active FROM ref_project_managers ORDER BY id"))
            add_custom_sheet("Ref - Project Statuses", _rows_to_dicts(conn, "SELECT id, name, color, active FROM ref_project_status ORDER BY id"))
            add_custom_sheet("Ref - Expense Categories", _rows_to_dicts(conn, "SELECT id, name, color_bg, color_fg, active FROM ref_expense_category ORDER BY id"))
            add_custom_sheet("Ref - Doc Categories", _rows_to_dicts(conn, "SELECT id, name, active FROM ref_doc_category ORDER BY id"))
            try:
                add_custom_sheet("Ref - LOB EBF Multipliers", _rows_to_dicts(conn, "SELECT lob, ebf_mult, updated_at FROM lob_ebf_settings"))
            except Exception as e_ebf:
                print(f"[export_full_portfolio_excel_bytes] lob_ebf_settings: {e_ebf}")

            try:
                add_custom_sheet("Users & Staff Roles", _rows_to_dicts(conn, "SELECT id, username, display_name, role, pm_name, created_at FROM users ORDER BY id"))
            except Exception as e_usr:
                print(f"[export_full_portfolio_excel_bytes] users: {e_usr}")

    except Exception as e:
        print(f"[export_full_portfolio_excel_bytes] Error: {e}")
        traceback.print_exc()

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def export_csv_bytes() -> bytes:
    """Export all projects as CSV."""
    projects = fetch_all_projects()
    out = io.StringIO()
    writer = csv.writer(out)
    writer.writerow([
        "id", "co_no", "project_name", "business_line", "project_manager",
        "client_name", "order_number", "order_issue_date", "order_value",
        "variation_amt", "revised_order_value", "project_status",
        "physical_completion", "billed_amount", "collection_till_date",
        "itd_cost", "baseline_eac", "rev_latest_eac", "project_type", "parent_project",
    ])
    for p in projects:
        writer.writerow([
            p.get("id"), p.get("co_no"), p.get("project_name"),
            p.get("business_line"), p.get("project_manager"), p.get("client_name"),
            p.get("order_number"), p.get("order_issue_date"),
            _n(p.get("order_value")), _n(p.get("variation_amt")),
            _n(p.get("revised_order_value")), p.get("project_status"),
            _n(p.get("physical_completion")), _n(p.get("billed_amount")),
            _n(p.get("collection_till_date")), _n(p.get("itd_cost")),
            _n(p.get("baseline_eac")), _n(p.get("rev_latest_eac")),
            p.get("project_type"), p.get("parent_project"),
        ])
    return out.getvalue().encode("utf-8-sig")


def export_pdf_bytes() -> bytes:
    """Export all projects as a PDF table."""
    projects = fetch_all_projects()
    rows = []
    total_ov = 0.0
    for p in projects:
        ov = _rev_ov(p)
        total_ov += ov
        rows.append([
            p.get("co_no") or "—",
            (p.get("project_name") or "—")[:40],
            p.get("project_manager") or "—",
            p.get("project_status") or "—",
            f"{ov:,.0f}",
            f"{_n(p.get('itd_cost')):,.0f}",
            f"{_n(p.get('collection_till_date')):,.0f}",
        ])
    return _build_table_pdf(
        "Project Portfolio Report",
        f"All projects — {len(projects)} records",
        ["CO.No", "Project Name", "PM", "Status", "Order Value (₹)", "ITD (₹)", "Collection (₹)"],
        rows,
        footer_label="Total Order Value",
        footer_value=total_ov,
    )


def export_data_entry_pdf_bytes() -> bytes:
    """PDF export specifically used by the Data Entry page."""
    return export_pdf_bytes()


# ── LOB summary exports ───────────────────────────────────────

def _build_lob_summary() -> tuple[list, list[list], float]:
    """Returns (headers, rows, total_ov) for LOB summary."""
    projects = fetch_all_projects()
    lob_list = get_lob_list()

    def _lob(p):
        val = (p.get("business_line") or p.get("hier_business_line") or "Other").strip()
        return normalize_lob_name(val)

    agg: dict = {}
    for p in projects:
        lb = _lob(p)
        d  = agg.setdefault(lb, {"count": 0, "ov": 0.0, "exp": 0.0, "col": 0.0})
        d["count"] += 1
        d["ov"]    += _rev_ov(p)
        d["exp"]   += _n(p.get("itd_cost"))
        d["col"]   += _n(p.get("collection_till_date"))

    ordered = []
    seen = set()
    for lb in (lob_list or []):
        norm = normalize_lob_name(lb)
        if norm and norm.lower() not in seen:
            seen.add(norm.lower())
            ordered.append(norm)
    for lb in agg.keys():
        if lb and lb.lower() not in seen:
            seen.add(lb.lower())
            ordered.append(lb)

    ordered = [lb for lb in ordered if agg.get(lb, {}).get("ov", 0) > 0 or agg.get(lb, {}).get("exp", 0) > 0]
    headers = ["Line of Business", "Projects", "Order Value (₹)", "Expenditure (₹)", "Collection (₹)", "Realization %"]
    rows = []
    total_ov = 0.0
    for lb in ordered:
        d = agg.get(lb, {})
        ov  = d.get("ov",  0.0)
        col = d.get("col", 0.0)
        total_ov += ov
        pct = f"{col / ov * 100:.1f}%" if ov else "—"
        rows.append([
            lb,
            str(d.get("count", 0)),
            f"{ov:,.0f}",
            f"{d.get('exp', 0.0):,.0f}",
            f"{col:,.0f}",
            pct,
        ])
    return headers, rows, total_ov


def export_lob_csv() -> bytes:
    headers, rows, _ = _build_lob_summary()
    out = io.StringIO()
    writer = csv.writer(out)
    writer.writerow(headers)
    writer.writerows(rows)
    return out.getvalue().encode("utf-8-sig")


def export_lob_xlsx() -> bytes:
    headers, rows, total_ov = _build_lob_summary()
    wb = Workbook()
    ws = wb.active
    ws.title = "LOB Summary"

    from openpyxl.styles import Font, PatternFill, Alignment
    hdr_fill = PatternFill("solid", fgColor="000E24")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)

    ws.append(headers)
    for cell in ws[1]:
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = Alignment(horizontal="center")

    for row in rows:
        ws.append(row)

    ws.append(["TOTAL", "", f"{total_ov:,.0f}", "", "", ""])
    last = ws.max_row
    for cell in ws[last]:
        cell.font = Font(bold=True, size=10)

    ws.column_dimensions["A"].width = 32
    for col in ["B", "C", "D", "E", "F"]:
        ws.column_dimensions[col].width = 18

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def export_lob_pdf() -> bytes:
    headers, rows, total_ov = _build_lob_summary()
    return _build_table_pdf(
        "LOB Performance Summary",
        f"Line-of-business breakdown — {sum(int(r[1]) for r in rows if r[1].isdigit())} projects",
        headers,
        rows,
        footer_label="Total Order Value",
        footer_value=total_ov,
    )

def _parse_body(handler) -> dict:
    """Parse POST body — handles both application/x-www-form-urlencoded
    and multipart/form-data (sent by fetch() with FormData objects).
    Returns a flat {field: value} dict (text fields only; files are skipped)."""
    import email as _email, email.policy as _email_policy

    content_type = handler.headers.get("Content-Type", "")
    n = int(handler.headers.get("Content-Length", 0))
    raw = handler.rfile.read(n)

    if "multipart/form-data" in content_type:
        # Build a fake MIME message so the standard email parser can decode it
        full_msg_bytes = (
            f"Content-Type: {content_type}\r\n\r\n".encode() + raw
        )
        try:
            msg = _email.message_from_bytes(
                full_msg_bytes, policy=_email_policy.compat32
            )
            result = {}
            for part in msg.walk():
                cd = part.get("Content-Disposition", "")
                if not cd or "filename=" in cd:
                    # skip the container part and any file-upload parts
                    continue
                # extract field name
                name = None
                for segment in cd.split(";"):
                    segment = segment.strip()
                    if segment.startswith("name="):
                        name = segment.split("=", 1)[1].strip().strip('"\'')
                        break
                if name is None:
                    continue
                payload = part.get_payload(decode=True)
                if payload is not None:
                    result[name] = payload.decode("utf-8", errors="replace")
            return result
        except Exception:
            return {}
    else:
        # application/x-www-form-urlencoded (default HTML form POST)
        decoded = raw.decode("utf-8", errors="replace")
        return {k: v[0] for k, v in parse_qs(decoded, keep_blank_values=True).items()}

def _drain_body(handler):
    try:
        n = int(handler.headers.get("Content-Length", 0))
        if n > 0: handler.rfile.read(n)
    except Exception:
        pass

_DEAD_CONN_ERRORS = (ConnectionAbortedError, ConnectionResetError, BrokenPipeError, OSError)

def _html_resp(handler, body: str, code=200):
    b = body.encode("utf-8")
    try:
        handler.send_response(code)
        handler.send_header("Content-Type",   "text/html; charset=utf-8")
        handler.send_header("Content-Length", str(len(b)))
        handler.end_headers()
        handler.wfile.write(b)
    except _DEAD_CONN_ERRORS:
        pass  # client disconnected before the response finished — nothing to do

def _redirect(handler, loc="/"):
    try:
        handler.send_response(302)
        handler.send_header("Location", loc)
        handler.end_headers()
    except _DEAD_CONN_ERRORS:
        pass

def _file_resp(handler, data: bytes, ctype: str, fname: str):
    try:
        handler.send_response(200)
        handler.send_header("Content-Type",        ctype)
        handler.send_header("Content-Disposition", f'attachment; filename="{fname}"')
        handler.send_header("Content-Length",      str(len(data)))
        handler.end_headers()
        handler.wfile.write(data)
    except _DEAD_CONN_ERRORS:
        pass

def _process_upload(tmp_path: str, suffix: str) -> str:
    """Import the complete Monarch portfolio workbook.

    The workbook is matched by sheet name and column name, never by column
    position.  Projects are the master table; Expenditures and Billing &
    Collection are linked back to projects using CO No.  The two report-style
    sheets (Sorting of Active RailwProject / Zero Collection) are not inserted
    as duplicate projects because their rows are summaries of the master data.
    """
    import openpyxl, math, json as _json_import

    def _norm_header(value):
        if value is None:
            return ""
        s = str(value).strip().lower()
        s = s.replace("&", "and")
        s = re.sub(r"[^a-z0-9]+", "_", s).strip("_")
        return s

    def _sheet_key(value):
        return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())

    def _safe_num(v):
        if v is None or isinstance(v, bool):
            return 0.0
        if isinstance(v, (int, float)):
            try:
                return 0.0 if math.isnan(float(v)) else float(v)
            except Exception:
                return 0.0
        s = str(v).strip()
        if not s or s.startswith("="):
            return 0.0
        s = s.replace(",", "").replace("₹", "").replace("%", "").strip()
        try:
            return float(s)
        except Exception:
            return 0.0

    def _safe_text(v):
        if v is None or isinstance(v, bool):
            return ""
        s = str(v).strip()
        if s.startswith("="):
            return ""
        return " ".join(s.split())

    def _safe_date(v):
        if v is None or isinstance(v, bool):
            return ""
        if hasattr(v, "strftime"):
            return v.strftime("%Y-%m-%d")
        s = str(v).strip()
        if not s or s.startswith("="):
            return ""
        for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y",
                    "%d.%m.%Y", "%d-%b-%Y", "%d/%b/%Y"):
            try:
                return datetime.strptime(s[:19], fmt).strftime("%Y-%m-%d")
            except Exception:
                pass
        return s[:10]

    def _percent(v):
        """Excel percentages may be 0.30 or 30; store as app percentage."""
        n = _safe_num(v)
        return n * 100.0 if 0 < abs(n) <= 1.0 else n

    def _headers_and_rows(ws):
        vals = list(ws.iter_rows(values_only=True))
        if not vals:
            return [], []
        # Find the first useful header row. Most portfolio sheets use row 1,
        # but accepting a few leading blank/title rows makes the importer safe.
        hdr_idx = 0
        for i, row in enumerate(vals[:10]):
            normalized = {_norm_header(c) for c in row if c is not None}
            if ("co_no" in normalized or "project_id" in normalized or
                    "id" in normalized or "project_name" in normalized):
                hdr_idx = i
                break
        headers = [_norm_header(c) for c in vals[hdr_idx]]
        return headers, vals[hdr_idx + 1:]

    def _index(headers, *names):
        wanted = {_norm_header(n) for n in names}
        for i, h in enumerate(headers):
            if h in wanted:
                return i
        return None

    def _get(row, idx):
        return row[idx] if idx is not None and idx < len(row) else None

    try:
        wb = openpyxl.load_workbook(tmp_path, read_only=True, data_only=True)
    except Exception as e:
        return f'<div class="alert err">❌ Could not open workbook: {html.escape(str(e))}</div>'

    # Locate sheets by normalized names so trailing spaces/case do not matter.
    sheets = {_sheet_key(ws.title): ws for ws in wb.worksheets}

    project_ws = sheets.get("projects")
    if project_ws is None:
        return '<div class="alert err">❌ The workbook must contain a "Projects" sheet.</div>'

    proj_headers, proj_rows = _headers_and_rows(project_ws)
    if not proj_headers:
        return '<div class="alert err">❌ The "Projects" sheet is empty.</div>'

    # Exact aliases for the 35 columns in the supplied portfolio workbook.
    pc = {
        "co_no": _index(proj_headers, "Co No", "Co.No", "CO.NO"),
        "project_name": _index(proj_headers, "Project Name"),
        "business_line": _index(proj_headers, "Business Line"),
        "project_manager": _index(proj_headers, "Project Manager", "PM"),
        "client_name": _index(proj_headers, "Client Name"),
        "order_number": _index(proj_headers, "Order Number"),
        "order_issue_date": _index(proj_headers, "Order Issue Date", "Order Date"),
        "order_value": _index(proj_headers, "Order Value/OA Amt", "Order Value", "OA Amt"),
        "revised_order_value": _index(proj_headers, "Revised Order Value"),
        "main_project_order_value": _index(proj_headers, "Main Project Order Value"),
        "hier_business_line": _index(proj_headers, "Hier Business Line", "LOB"),
        "parent_project": _index(proj_headers, "Parent Project"),
        "child_project": _index(proj_headers, "Child Project"),
        "baseline_eac": _index(proj_headers, "Baseline Eac", "Baseline EAC"),
        "baseline_margin": _index(proj_headers, "Baseline Margin"),
        "rev_latest_eac": _index(proj_headers, "Rev Latest Eac", "Latest EAC"),
        "rev_latest_margin": _index(proj_headers, "Rev Latest Margin", "Latest Margin"),
        "project_status": _index(proj_headers, "Project Status"),
        "physical_completion": _index(proj_headers, "Physical Completion"),
        "billed_amount": _index(proj_headers, "Uptodate Bill Amt", "Billed Amount"),
        "collection_till_date": _index(proj_headers, "Collection till date", "Collection Till Date"),
        "itd_cost": _index(proj_headers, "Total ITD", "Itd Cost", "ITD Cost", "ITD Expenditure", "Total ITD Cost", "Total Expenses", "ITD"),
        "project_type": _index(proj_headers, "Project Type"),
        "parent_id": _index(proj_headers, "Parent Id", "Parent ID", "Parent Project Co.No"),
        "created_at": _index(proj_headers, "Created At"),
        "variation_amt": _index(proj_headers, "Variation Amt", "Variation"),
        "last_billed_date": _index(proj_headers, "Last Billed Date"),
        "bill_submission_pending_date": _index(
            proj_headers, "Bill Submission Pending Date",
            "Bill submission date but not yet collected"
        ),
        "folder_path": _index(proj_headers, "Folder Path"),
        "project_duration": _index(proj_headers, "Project Duration", "Duration in Months"),
        "duration_unit": _index(proj_headers, "Duration Unit"),
        "end_date": _index(proj_headers, "End Date"),
        "extension_months": _index(proj_headers, "Extension Months"),
        "extended_end_date": _index(proj_headers, "Extended End Date"),
        "sub_project": _index(proj_headers, "Sub Project"),
    }

    if pc["itd_cost"] is None and len(proj_headers) >= 21:
        if any(term in proj_headers[20] for term in ("itd", "cost", "exp")):
            pc["itd_cost"] = 20

    missing_required = [k for k in ("co_no", "project_name") if pc[k] is None]
    if missing_required:
        return (f'<div class="alert err">❌ Projects sheet is missing required '
                f'column(s): {", ".join(missing_required)}.</div>')

    # Read all project rows first and build an in-memory authoritative mapping.
    projects_in = []
    seen_co = set()
    duplicate_co = []
    for row_num, row in enumerate(proj_rows, start=2):
        if not row or all(c is None for c in row):
            continue
        co_no = _safe_text(_get(row, pc["co_no"]))
        name = _safe_text(_get(row, pc["project_name"]))
        if not co_no:
            continue
        if co_no in seen_co:
            duplicate_co.append(f"row {row_num}: {co_no}")
            continue
        seen_co.add(co_no)

        ov = _safe_num(_get(row, pc["order_value"]))
        rev = _safe_num(_get(row, pc["revised_order_value"]))
        var_amt = _safe_num(_get(row, pc["variation_amt"]))
        if not rev:
            rev = ov + var_amt

        parent_val = _safe_text(_get(row, pc["parent_project"]))
        parent_id = _safe_text(_get(row, pc["parent_id"]))
        project_type = _safe_text(_get(row, pc["project_type"])).lower()
        child_project = _safe_text(_get(row, pc["child_project"])) or _safe_text(_get(row, pc["sub_project"]))

        # Parent ID is authoritative when present. "0" is treated as blank.
        if parent_id in ("0", "0.0"):
            parent_id = ""

        # Physical completion and margins are stored in the DB as percentages.
        phys = _percent(_get(row, pc["physical_completion"]))
        b_margin = _percent(_get(row, pc["baseline_margin"]))
        rl_margin = _percent(_get(row, pc["rev_latest_margin"]))

        # Preserve the workbook's explicit End Date. If it is absent, derive it
        # from Order Issue Date + Project Duration when possible.
        order_date = _safe_date(_get(row, pc["order_issue_date"]))
        end_date = _safe_date(_get(row, pc["end_date"]))
        duration = _safe_num(_get(row, pc["project_duration"]))
        if not end_date and order_date and duration:
            try:
                end_date = _add_months(datetime.fromisoformat(order_date).date(),
                                       int(round(duration))).isoformat()
            except Exception:
                end_date = ""

        if parent_id or parent_val:
            project_type = "child"
        elif project_type not in ("child", "sub", "sub project"):
            project_type = "parent"
        else:
            project_type = "child"

        projects_in.append({
            "co_no": co_no,
            "project_name": name,
            "business_line": normalize_lob_name(_safe_text(_get(row, pc["business_line"]))),
            "project_manager": _safe_text(_get(row, pc["project_manager"])),
            "client_name": _safe_text(_get(row, pc["client_name"])),
            "order_number": _safe_text(_get(row, pc["order_number"])),
            "order_issue_date": order_date,
            "end_date": end_date,
            "extension_months": _safe_num(_get(row, pc["extension_months"])),
            "extended_end_date": _safe_date(_get(row, pc["extended_end_date"])),
            "order_value": ov,
            "variation_amt": var_amt,
            "revised_order_value": rev,
            "hier_business_line": normalize_lob_name(_safe_text(_get(row, pc["hier_business_line"]))),
            "parent_project": parent_val,
            "parent_id": parent_id,
            "child_project": child_project,
            "baseline_eac": _safe_num(_get(row, pc["baseline_eac"])),
            "baseline_margin": b_margin,
            "rev_latest_eac": _safe_num(_get(row, pc["rev_latest_eac"])),
            "rev_latest_margin": rl_margin,
            "project_status": _safe_text(_get(row, pc["project_status"])) or "Active",
            "physical_completion": phys,
            "billed_amount": _safe_num(_get(row, pc["billed_amount"])),
            "last_billed_date": _safe_date(_get(row, pc["last_billed_date"])),
            "collection_till_date": _safe_num(_get(row, pc["collection_till_date"])),
            "itd_cost": _safe_num(_get(row, pc["itd_cost"])),
            "bill_submission_pending_date": _safe_date(_get(row, pc["bill_submission_pending_date"])),
            "project_type": project_type,
            "folder_path": _safe_text(_get(row, pc["folder_path"])),
            "created_at": _safe_text(_get(row, pc["created_at"])) or datetime.now().isoformat(),
        })

    if not projects_in:
        return '<div class="alert err">❌ No project records were found in the Projects sheet.</div>'

    # Read Expenditures. The supplied workbook has 19 columns; all cost columns
    # are mapped by name and linked with CO No rather than the old numeric ID.
    exp_in = []
    exp_ws = sheets.get("expenditures")
    if exp_ws:
        eh, erows = _headers_and_rows(exp_ws)
        ec = {
            "co_no": _index(eh, "Co No", "CO.No"),
            "project_id": _index(eh, "Project Id", "Project ID"),
            "lob": _index(eh, "Lob", "LOB"),
            "data_as_of": _index(eh, "Data As Of"),
            "erp_update_date": _index(eh, "Erp Update Date", "ERP Update Date"),
            "itd_cogs": _index(eh, "Itd Cogs", "ITD COGS"),
            "itd_djcs": _index(eh, "Itd Djcs", "ITD DJCS"),
            "itd_equip": _index(eh, "Itd Equip", "ITD Equipment"),
            "itd_manpower": _index(eh, "Itd Manpower"),
            "itd_vehicles": _index(eh, "Itd Vehicles"),
            "itd_others": _index(eh, "Itd Others"),
            "erp_cogs": _index(eh, "Erp Cogs", "ERP COGS"),
            "erp_djcs": _index(eh, "Erp Djcs", "ERP DJCS"),
            "erp_equip": _index(eh, "Erp Equip", "ERP Equipment"),
            "erp_manpower": _index(eh, "Erp Manpower"),
            "erp_vehicles": _index(eh, "Erp Vehicles"),
            "erp_others": _index(eh, "Erp Others"),
            "created_at": _index(eh, "Created At"),
        }
        for row in erows:
            if not row or all(c is None for c in row):
                continue
            co = _safe_text(_get(row, ec["co_no"]))
            if not co:
                # Fallback only when CO No is absent: map old Project Id to the
                # workbook's project row by ordinal/id where possible.
                raw_pid = _safe_text(_get(row, ec["project_id"]))
                try:
                    idx = int(float(raw_pid)) - 1
                    if 0 <= idx < len(projects_in):
                        co = projects_in[idx]["co_no"]
                except Exception:
                    pass
            if not co:
                continue
            exp_in.append({
                "co_no": co,
                "lob": _safe_text(_get(row, ec["lob"])),
                "data_as_of": _safe_date(_get(row, ec["data_as_of"])),
                "erp_update_date": _safe_date(_get(row, ec["erp_update_date"])),
                "itd_cogs": _safe_num(_get(row, ec["itd_cogs"])),
                "itd_djcs": _safe_num(_get(row, ec["itd_djcs"])),
                "itd_equip": _safe_num(_get(row, ec["itd_equip"])),
                "itd_manpower": _safe_num(_get(row, ec["itd_manpower"])),
                "itd_vehicles": _safe_num(_get(row, ec["itd_vehicles"])),
                "itd_others": _safe_num(_get(row, ec["itd_others"])),
                "erp_cogs": _safe_num(_get(row, ec["erp_cogs"])),
                "erp_djcs": _safe_num(_get(row, ec["erp_djcs"])),
                "erp_equip": _safe_num(_get(row, ec["erp_equip"])),
                "erp_manpower": _safe_num(_get(row, ec["erp_manpower"])),
                "erp_vehicles": _safe_num(_get(row, ec["erp_vehicles"])),
                "erp_others": _safe_num(_get(row, ec["erp_others"])),
                "created_at": _safe_text(_get(row, ec["created_at"])) or datetime.now().isoformat(),
            })

    # Read Billing & Collection. Multiple ledger rows per project are preserved.
    billing_in = []
    bc_ws = sheets.get("billingcollection")
    if bc_ws:
        bh, brows = _headers_and_rows(bc_ws)
        bc = {
            "project_id": _index(bh, "Project Id", "Project ID"),
            "ra_no": _index(bh, "Ra No", "RA No", "RA No."),
            "billing_date": _index(bh, "Billing Date"),
            "billed_amt": _index(bh, "Billed Amt", "Billed Amount"),
            "collection_date": _index(bh, "Collection Date"),
            "collection_amt": _index(bh, "Collection Amt", "Collection Amount"),
            "created_at": _index(bh, "Created At"),
            "project_name": _index(bh, "Project Name"),
            "co_no": _index(bh, "Co No", "CO.No"),
        }
        for row in brows:
            if not row or all(c is None for c in row):
                continue
            co = _safe_text(_get(row, bc["co_no"]))
            if not co:
                pname = _safe_text(_get(row, bc["project_name"]))
                if pname:
                    candidates = [p["co_no"] for p in projects_in
                                  if p["project_name"].strip().lower() == pname.strip().lower()]
                    if len(candidates) == 1:
                        co = candidates[0]
            if not co:
                continue
            billing_in.append({
                "co_no": co,
                "ra_no": _safe_text(_get(row, bc["ra_no"])),
                "billing_date": _safe_date(_get(row, bc["billing_date"])),
                "billed_amt": _safe_num(_get(row, bc["billed_amt"])),
                "collection_date": _safe_date(_get(row, bc["collection_date"])),
                "collection_amt": _safe_num(_get(row, bc["collection_amt"])),
                "created_at": _safe_text(_get(row, bc["created_at"])) or datetime.now().isoformat(),
            })

    # The report sheets are intentionally not duplicated into transactional
    # tables. Their underlying project/expenditure data is imported above.
    report_sheet_notes = []
    for key, label in (
        ("sortingofactiverailwproject", "Sorting of Active RailwProject"),
        ("zerocollection", "Zero Collection"),
    ):
        if key in sheets:
            report_sheet_notes.append(label)

    # Validate all linked rows before destructive replacement.
    project_cos = {p["co_no"] for p in projects_in}
    bad_exp = [e["co_no"] for e in exp_in if e["co_no"] not in project_cos]
    bad_bc = [b["co_no"] for b in billing_in if b["co_no"] not in project_cos]
    if bad_exp or bad_bc:
        details = []
        if bad_exp:
            details.append(f"{len(bad_exp)} expenditure row(s) have an unknown CO No")
        if bad_bc:
            details.append(f"{len(bad_bc)} billing row(s) have an unknown CO No")
        return '<div class="alert err">❌ Upload stopped — ' + "; ".join(details) + \
               '. No existing database data was deleted.</div>'

    try:
        # Full replacement happens only after the workbook has passed validation.
        wipe_all_transactional_data()

        co_to_dbid = {}
        inserted_projects = 0
        linked_children = 0

        # Insert parents first so child parent_id relationships can be resolved.
        ordered = sorted(projects_in, key=lambda p: 1 if p["project_type"] == "child" else 0)

        for p in ordered:
            parent_co = p["parent_id"]
            if parent_co and parent_co not in project_cos:
                # Some workbooks put the parent project name instead of CO No.
                match = next((x["co_no"] for x in projects_in
                              if x["project_name"].strip().lower() == parent_co.strip().lower()), "")
                parent_co = match

            parent_name = p["parent_project"]
            if parent_co:
                parent_row = next((x for x in projects_in if x["co_no"] == parent_co), None)
                if parent_row:
                    parent_name = parent_row["project_name"]

            # Keep Excel's child/parent classification, but normalize parent fields.
            project_type = "child" if (p["project_type"] == "child" or parent_co) else "parent"
            if project_type == "parent":
                parent_co = None
                parent_name = ""

            fields = dict(p)
            fields["parent_id"] = parent_co
            fields["parent_project"] = parent_name
            fields["project_type"] = project_type

            # Use direct SQL rather than save_project() so the imported Excel
            # values (including margins and created_at) are not silently replaced
            # by the manual-form calculations.
            with ENGINE.begin() as conn:
                row = conn.execute(text("""
                    INSERT INTO projects (
                        co_no, project_name, business_line, project_manager, client_name,
                        order_number, order_issue_date, end_date, extension_months,
                        extended_end_date, order_value, variation_amt, revised_order_value,
                        hier_business_line, parent_project, parent_id, child_project,
                        baseline_eac, baseline_margin, rev_latest_eac, rev_latest_margin,
                        project_status, physical_completion, billed_amount, last_billed_date,
                        collection_till_date, itd_cost, bill_submission_pending_date,
                        project_type, folder_path, created_at
                    ) VALUES (
                        :co_no, :project_name, :business_line, :project_manager, :client_name,
                        :order_number, :order_issue_date, :end_date, :extension_months,
                        :extended_end_date, :order_value, :variation_amt, :revised_order_value,
                        :hier_business_line, :parent_project, :parent_id, :child_project,
                        :baseline_eac, :baseline_margin, :rev_latest_eac, :rev_latest_margin,
                        :project_status, :physical_completion, :billed_amount, :last_billed_date,
                        :collection_till_date, :itd_cost, :bill_submission_pending_date,
                        :project_type, :folder_path, :created_at
                    ) RETURNING id
                """), fields)
                dbid = row.fetchone()[0]

            co_to_dbid[p["co_no"]] = dbid
            inserted_projects += 1

            for ref_table, ref_val in (
                ("ref_project_managers", p["project_manager"]),
                ("ref_clients", p["client_name"]),
                ("ref_lob", normalize_lob_name(p["business_line"] or p["hier_business_line"])),
                ("ref_business_line", normalize_lob_name(p["business_line"])),
            ):
                if ref_val:
                    try:
                        add_ref_value(ref_table, "name", ref_val)
                    except Exception:
                        pass

        # Rebuild parent child_project text from the imported relationships.
        with ENGINE.begin() as conn:
            for p in projects_in:
                if p["project_type"] != "child":
                    continue
                parent_co = p["parent_id"]
                if parent_co in co_to_dbid:
                    linked_children += 1
                    parent_name = p["project_name"]
                    conn.execute(text("""
                        UPDATE projects
                        SET child_project = CASE
                            WHEN child_project IS NULL OR child_project = '' THEN :child
                            WHEN child_project NOT LIKE '%' || :child || '%' THEN child_project || ', ' || :child
                            ELSE child_project
                        END
                        WHERE co_no=:pco AND project_type='parent'
                    """), {"child": parent_name, "pco": parent_co})

        # Insert expenditures and create one ERP history row per imported ERP head.
        # This makes the existing ERP reconciliation UI show the imported totals.
        erp_map = (
            ("erp_cogs", "COGS"),
            ("erp_djcs", "DJCS"),
            ("erp_equip", "Equipment"),
            ("erp_manpower", "Manpower"),
            ("erp_vehicles", "Vehicles"),
            ("erp_others", "Others"),
        )
        proj_itd_by_co = {p["co_no"]: _safe_num(p.get("itd_cost")) for p in projects_in}
        inserted_exp = 0
        touched_exp_projects = set()
        for e in exp_in:
            pid = co_to_dbid[e["co_no"]]
            co = e["co_no"]
            p_itd = proj_itd_by_co.get(co, 0.0)
            erp_sum = sum(_safe_num(e[k]) for k, _ in erp_map)
            diff = p_itd - erp_sum
            if diff > 0.001:
                e["erp_others"] = _safe_num(e.get("erp_others")) + diff
                e["itd_others"] = _safe_num(e.get("itd_others")) + diff

            with ENGINE.begin() as conn:
                r = conn.execute(text("""
                    INSERT INTO expenditures (
                        project_id, co_no, lob, data_as_of, erp_update_date,
                        itd_cogs, itd_djcs, itd_equip, itd_manpower, itd_vehicles, itd_others,
                        erp_cogs, erp_djcs, erp_equip, erp_manpower, erp_vehicles, erp_others,
                        created_at
                    ) VALUES (
                        :project_id, :co_no, :lob, :data_as_of, :erp_update_date,
                        :itd_cogs, :itd_djcs, :itd_equip, :itd_manpower, :itd_vehicles, :itd_others,
                        :erp_cogs, :erp_djcs, :erp_equip, :erp_manpower, :erp_vehicles, :erp_others,
                        :created_at
                    ) RETURNING id
                """), {**e, "project_id": pid})
                exp_id = r.fetchone()[0]
                inserted_exp += 1

                for key, category in erp_map:
                    amount = _safe_num(e[key])
                    if amount:
                        conn.execute(text("""
                            INSERT INTO erp_updates
                                (project_id, co_no, category, amount, created_by, created_at)
                            VALUES (:pid, :co, :cat, :amt, :by, :ts)
                        """), {
                            "pid": pid, "co": e["co_no"], "cat": category,
                            "amt": amount, "by": "excel_upload",
                            "ts": e["created_at"],
                        })
            touched_exp_projects.add(pid)

        # For projects that were in the Projects sheet with itd_cost > 0 but NOT in the Expenditures sheet,
        # create an expenditures record and erp_updates record under "Others" so they are tracked in expenditures.
        exp_cos = {e["co_no"] for e in exp_in}
        today_iso = datetime.now().date().isoformat()
        for p in projects_in:
            co = p["co_no"]
            if co not in exp_cos and p["itd_cost"] > 0 and co in co_to_dbid:
                pid = co_to_dbid[co]
                with ENGINE.begin() as conn:
                    conn.execute(text("""
                        INSERT INTO expenditures (
                            project_id, co_no, lob, data_as_of, erp_update_date,
                            itd_cogs, itd_djcs, itd_equip, itd_manpower, itd_vehicles, itd_others,
                            erp_cogs, erp_djcs, erp_equip, erp_manpower, erp_vehicles, erp_others,
                            created_at
                        ) VALUES (
                            :project_id, :co_no, :lob, :today, :today,
                            0, 0, 0, 0, 0, :itd,
                            0, 0, 0, 0, 0, :itd,
                            :created_at
                        )
                    """), {
                        "project_id": pid, "co_no": co, "lob": p.get("hier_business_line") or p.get("business_line") or "",
                        "today": today_iso, "itd": p["itd_cost"],
                        "created_at": p.get("created_at") or datetime.now().isoformat()
                    })
                    inserted_exp += 1
                    conn.execute(text("""
                        INSERT INTO erp_updates
                            (project_id, co_no, category, amount, created_by, created_at)
                        VALUES (:pid, :co, 'Others', :amt, 'excel_upload', :ts)
                    """), {
                        "pid": pid, "co": co, "amt": p["itd_cost"],
                        "ts": p.get("created_at") or datetime.now().isoformat()
                    })

        # ── Ensure projects.itd_cost preserves the Projects sheet's authoritative
        #    ITD Cost (e.g. Column U). Only backfill from erp_updates if itd_cost
        #    was 0 or not provided in the Projects sheet. ──
        for pid in touched_exp_projects:
            try:
                with ENGINE.begin() as conn:
                    current_itd = conn.execute(text(
                        "SELECT itd_cost FROM projects WHERE id=:id"
                    ), {"id": pid}).fetchone()
                    if not current_itd or not current_itd[0]:
                        _sync_project_itd_cost(pid)
            except Exception:
                pass

        # Billing ledger: preserve every row, then sync project totals from the ledger.
        inserted_bc = 0
        touched_projects = set()
        for b in billing_in:
            pid = co_to_dbid[b["co_no"]]
            with ENGINE.begin() as conn:
                conn.execute(text("""
                    INSERT INTO billing_collection
                        (project_id, ra_no, billing_date, billed_amt,
                         collection_date, collection_amt, created_at)
                    VALUES (:project_id, :ra_no, :billing_date, :billed_amt,
                            :collection_date, :collection_amt, :created_at)
                """), {**b, "project_id": pid})
            inserted_bc += 1
            touched_projects.add(pid)

        for pid in touched_projects:
            try:
                _sync_project_billing_totals(pid)
            except Exception:
                pass

        # Register report-sheet LOB values as well, without creating duplicate projects.
        for report_key in ("sortingofactiverailwproject", "zerocollection"):
            ws = sheets.get(report_key)
            if ws:
                rh, rr = _headers_and_rows(ws)
                li = _index(rh, "LOB", "Hier Business Line", "Business Line")
                if li is not None:
                    for row in rr:
                        val = _safe_text(_get(row, li))
                        if val:
                            try:
                                add_ref_value("ref_lob", "name", val)
                            except Exception:
                                pass

        summary = (
            f"✅ Portfolio upload complete — {inserted_projects:,} projects, "
            f"{inserted_exp:,} expenditure records, {inserted_bc:,} billing/collection records "
            f"and {linked_children:,} child-project links mapped."
        )
        extra = []
        if report_sheet_notes:
            extra.append(
                "Report sheets kept as summaries (not duplicated): " +
                ", ".join(report_sheet_notes) + "."
            )
        if duplicate_co:
            extra.append(
                f"{len(duplicate_co)} duplicate CO No row(s) were skipped, keeping the first occurrence."
            )
        if extra:
            summary += " " + " ".join(extra)
        return f'<div class="alert ok">{html.escape(summary)}</div>'

    except Exception as e:
        return (f'<div class="alert err">❌ Upload failed: {html.escape(str(e))}<br>'
                f'<pre style="font-size:11px">{html.escape(traceback.format_exc())}</pre></div>')

class MonarchHandler(SimpleHTTPRequestHandler):
    def log_message(self, fmt, *args): pass

    def do_GET(self):
        path = urlparse(self.path).path.rstrip("/") or "/"
        _req_ctx.path = self.path
        try:
            if path == "/login":
                _html_resp(self, login_page()); return
            if path == "/reset_password":
                _html_resp(self, reset_password_page()); return
            if path == "/logout":
                token = _get_session_token(self)
                if token: destroy_session(token)
                self.send_response(302)
                self.send_header("Set-Cookie",
                    f"{SESSION_COOKIE}=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax")
                self.send_header("Location", "/login")
                self.end_headers(); return
            if path == "/logo.jpg":
                data = LOGO_BYTES
                self.send_response(200)
                self.send_header("Content-Type", "image/jpeg")
                self.send_header("Content-Length", str(len(data)))
                self.send_header("Cache-Control", "public, max-age=86400")
                self.end_headers()
                self.wfile.write(data)
                return

            user = get_session_user(self)
            if not user: _redirect(self, "/login"); return
            role = get_session_role(self)
            _req_ctx.role     = role
            _req_ctx.is_admin = is_admin_role(role)
            _req_ctx.pm_name  = get_session_pm_name(self)
            _req_ctx.active_only = get_session_active_filter(self)

            if path == "/api/data_version":
                import json as _json
                payload = _json.dumps({"v": get_data_version()}).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type",   "application/json")
                self.send_header("Cache-Control",  "no-store")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return

            if role == "director" and path not in DIRECTOR_ALLOWED_PATHS:
                _redirect(self, "/variations"); return

            if role == "accounts" and path not in ACCOUNTS_ALLOWED_PATHS:
                _redirect(self, "/payment_approvals"); return

            if role == "procurement" and path not in PROCUREMENT_ALLOWED_PATHS:
                _redirect(self, "/payment_approvals"); return

            if role in ("pm", "head") and path in PM_BLOCKED_PATHS:
                _redirect(self, "/dashboard"); return

            if path == "/toggle_active_filter":
                toggle_session_active_filter(self)
                qs  = parse_qs(urlparse(self.path).query)
                nxt = (qs.get("next", ["/dashboard"])[0] or "/dashboard").strip()
                if not nxt.startswith("/"):
                    nxt = "/dashboard"
                _redirect(self, nxt); return

            if   path in ("/","/dashboard","/index"): _html_resp(self, dashboard(user, _req_ctx.is_admin))
            elif path == "/data_entry":
                qs         = parse_qs(urlparse(self.path).query)
                pid_qs     = qs.get("id", [""])[0].strip()
                status_qs  = qs.get("status", [""])[0].strip()
                if pid_qs and pid_qs.isdigit():
                    proj = fetch_project_by_id(int(pid_qs))
                    if proj:
                        _html_resp(self, project_data_entry(prefill=proj, username=user, status_filter=status_qs)); return
                _html_resp(self, project_data_entry(username=user, status_filter=status_qs))
            elif path == "/search":
                q = parse_qs(urlparse(self.path).query).get("q", [""])[0]
                _html_resp(self, search_results_page(q, username=user))
            elif path == "/expenditures":
                qs = parse_qs(urlparse(self.path).query)
                prefill = {k: v[0] for k, v in qs.items()} if qs else None
                _html_resp(self, expenditures(prefill=prefill, username=user))
            elif path == "/financials":
                qs = parse_qs(urlparse(self.path).query)
                prefill = {}
                if qs.get("lob"): prefill["lob"] = qs.get("lob", [""])[0]
                if qs.get("pm"):  prefill["pm"]  = qs.get("pm", [""])[0]
                _html_resp(self, financials(prefill=prefill or None, username=user))
            elif path == "/chat":         _html_resp(self, chat_page(username=user))
            elif path == "/variations":   _html_resp(self, variations(username=user))
            elif path == "/reports":
                section = parse_qs(urlparse(self.path).query).get("section", [""])[0]
                _html_resp(self, reports(username=user, active_section=section))
            elif path == "/kpi":          _html_resp(self, kpi(username=user))
            elif path == "/eac_status":
                qs = parse_qs(urlparse(self.path).query)
                _html_resp(self, eac_status_report(
                    username=user,
                    pm=qs.get("pm", [""])[0],
                    status=qs.get("status", [""])[0],
                    q=qs.get("q", [""])[0],
                    page=int(qs.get("page", ["1"])[0] or 1),
                ))
            elif path == "/schedule":
                qs = parse_qs(urlparse(self.path).query)
                _html_resp(self, schedule_performance_report(
                    username=user,
                    lob=qs.get("lob", [""])[0],
                    pm=qs.get("pm", [""])[0],
                    status=qs.get("status", [""])[0],
                    q=qs.get("q", [""])[0],
                    page=int(qs.get("page", ["1"])[0] or 1),
                ))
            elif path == "/billing_collection":
                qs = parse_qs(urlparse(self.path).query)
                prefill = {k: v[0] for k, v in qs.items()} if qs else None
                _html_resp(self, billing_collection(prefill=prefill, username=user))
            elif path == "/payments_dashboard":
                _html_resp(self, payments_dashboard(username=user, role=role))
            elif path == "/weekly_booking":
                qs = parse_qs(urlparse(self.path).query)
                edit_id = qs.get("edit_id", [""])[0]
                edit_booking = None
                if edit_id.isdigit():
                    cand = fetch_booking_detail(int(edit_id))
                    if (cand and (cand.get("pm_username") or "").strip().lower() == user.strip().lower()
                            and (cand.get("overall_status") or "").lower() == "disapproved"):
                        edit_booking = cand
                _html_resp(self, weekly_payment_booking(
                    username=user, role=role, pm_name=getattr(_req_ctx, "pm_name", ""),
                    sel_project_id=qs.get("project_id", [str(edit_booking.get("project_id") or "") if edit_booking else ""])[0],
                    is_indirect=(qs.get("is_indirect", ["1" if edit_booking and edit_booking.get("is_indirect") else "0"])[0] == "1"),
                    edit_booking=edit_booking,
                ))
            elif path == "/payment_approvals":
                qs = parse_qs(urlparse(self.path).query)
                bid = qs.get("id", [""])[0]
                _html_resp(self, payment_approvals(
                    username=user, role=role,
                    selected_id=int(bid) if bid.isdigit() else None,
                ))
            elif path == "/export_excel/payment_bookings":
                _file_resp(self, export_payment_bookings_excel_bytes(), XLSX_MIME, "payment_bookings.xlsx")
            elif path == "/download_stamped_doc":
                qs = parse_qs(urlparse(self.path).query)
                bid = int(qs.get("id", ["0"])[0] or 0)
                if bid:
                    _file_resp(self, stamp_booking_pdf_bytes(bid), "application/pdf",
                               f"payment_booking_{bid}_stamped.pdf")
                else:
                    _redirect(self, "/payment_approvals")
            elif path == "/edit_project":
                pid  = int(parse_qs(urlparse(self.path).query).get("id",["0"])[0])
                proj = fetch_project_by_id(pid)
                if proj and role == "pm":
                    pm_name = (getattr(_req_ctx, "pm_name", "") or "").strip().lower()
                    if (proj.get("project_manager") or "").strip().lower() != pm_name:
                        proj = {}
                if proj:
                    _html_resp(self, project_data_entry(
                        alert_html=f'<div class="alert ok">✏ Editing project #{pid} — resubmit to save changes.</div>',
                        prefill=proj, username=user))
                else:
                    _redirect(self, "/data_entry")
            
            elif path == "/user_management":
              role = get_session_role(self)
              if not is_admin_role(role):
                  _html_resp(self,
                      _html_head("Access Denied") + "<body>" + _sidebar(is_admin=False) + _topbar(username=user) +
                      f'<div class="main"><div class="canvas" style="padding-top:60px;text-align:center">'
                      f'<p style="font-size:20px;color:{C["error"]};font-weight:700">🔒 Access Denied</p>'
                      f'<p style="color:#6b7280">You need administrator privileges to view this page.</p>'
                      f'<p><a href="/dashboard">← Back to Dashboard</a></p></div></div></body></html>',
                      code=403)
              else:
                  _html_resp(self, user_management(username=user))
            
            elif path == "/open_folder":
              import subprocess, urllib.parse, os, time, ctypes
              from ctypes import wintypes
              qs       = parse_qs(urlparse(self.path).query)
              raw_path = qs.get("path", [""])[0]
              folder   = urllib.parse.unquote(raw_path).replace("/", "\\")
              user     = get_session_user(self)
              if not user:
                  _redirect(self, "/login"); return
              try:
                  if folder and os.path.exists(folder):

                      def open_and_focus():
                          user32  = ctypes.windll.user32
                          kernel32= ctypes.windll.kernel32

                          # Open explorer
                          subprocess.Popen(["explorer.exe", folder])
                          time.sleep(0.8)

                          folder_name = os.path.basename(folder.rstrip("\\"))
                          found = []

                          def enum_cb(hwnd, _):
                              if user32.IsWindowVisible(hwnd):
                                  buf = ctypes.create_unicode_buffer(512)
                                  user32.GetWindowTextW(hwnd, buf, 512)
                                  if folder_name.lower() in buf.value.lower():
                                      found.append(hwnd)
                              return True

                          WNDENUMPROC = ctypes.WINFUNCTYPE(
                              ctypes.c_bool, wintypes.HWND, wintypes.LPARAM
                          )
                          user32.EnumWindows(WNDENUMPROC(enum_cb), 0)

                          if found:
                              hwnd = found[0]
                              # KEY TRICK: fake Alt keypress to allow SetForegroundWindow
                              VK_MENU = 0x12
                              KEYEVENTF_KEYUP = 0x0002
                              user32.keybd_event(VK_MENU, 0, 0, 0)
                              user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)
                              time.sleep(0.05)

                              # Force window to front
                              user32.ShowWindow(hwnd, 9)        # SW_RESTORE
                              user32.BringWindowToTop(hwnd)
                              user32.SetForegroundWindow(hwnd)
                              user32.SetActiveWindow(hwnd)
                              user32.SetFocus(hwnd)

                      import threading
                      threading.Thread(target=open_and_focus, daemon=True).start()

                      self.send_response(200)
                      self.send_header("Content-Type", "application/json")
                      self.end_headers()
                      self.wfile.write(b'{"ok":true}')
                  else:
                      self.send_response(404)
                      self.send_header("Content-Type", "application/json")
                      self.end_headers()
                      self.wfile.write(b'{"ok":false,"error":"Path not found"}')
              except Exception as e:
                  self.send_response(500)
                  self.send_header("Content-Type", "application/json")
                  self.end_headers()
                  self.wfile.write(f'{{"ok":false,"error":"{e}"}}'.encode())
              return

            elif path == "/export_excel": _file_resp(self, export_excel_bytes(),  XLSX_MIME, "monarch_projects.xlsx")
            elif path == "/export_excel/full":
                _file_resp(self, export_full_portfolio_excel_bytes(), XLSX_MIME,
                           f"monarch_full_portfolio_{date.today().isoformat()}.xlsx")
            elif path == "/export_csv":   _file_resp(self, export_csv_bytes(),    "text/csv","monarch_projects.csv")
            elif path == "/export_pdf":   _file_resp(self, export_pdf_bytes(),    "application/pdf","monarch_projects.pdf")
            elif path == "/export/csv":   _file_resp(self, export_lob_csv(),      "text/csv","lob_summary.csv")
            elif path == "/export/xlsx":  _file_resp(self, export_lob_xlsx(),     XLSX_MIME, "lob_summary.xlsx")
            elif path == "/export/pdf":   _file_resp(self, export_lob_pdf(),      "application/pdf","lob_summary.pdf")
            elif path == "/export_pdf/kpi":         _file_resp(self, export_pdf_bytes(),        "application/pdf", "kpi_report.pdf")
            elif path == "/export_pdf/dashboard":   _file_resp(self, export_dashboard_pdf_bytes(), "application/pdf", "dashboard_report.pdf")
            elif path == "/export_pdf/data_entry":  _file_resp(self, export_data_entry_pdf_bytes(), "application/pdf", "projects_report.pdf")
            elif path == "/export_pdf/financials":
                qs       = parse_qs(urlparse(self.path).query)
                lob      = qs.get("lob", [""])[0]
                ebf_mult = qs.get("ebf_mult", ["1.35"])[0]
                pm_sel   = qs.get("pm", [""])[0]
                _file_resp(self, export_financials_pdf_bytes(lob, ebf_mult, pm_sel), "application/pdf", "financials_report.pdf")
            elif path == "/export_pdf/variations":  _file_resp(self, export_variations_pdf_bytes(), "application/pdf", "variations_report.pdf")
            elif path == "/export_excel/reports":
                qs   = parse_qs(urlparse(self.path).query)
                sec  = qs.get("section", ["oi"])[0]
                lob  = qs.get("lob", [""])[0]
                q    = qs.get("q", [""])[0]
                fname = "consolidated_report.xlsx" if "," in sec else {
                    "oi": "order_intake_report.xlsx", "col": "collection_report.xlsx",
                    "itd": "billed_amount_report.xlsx", "ub": "unbilled_amount_report.xlsx",
                    "fin": "financial_pct_report.xlsx", "phy": "physical_pct_report.xlsx",
                    "dur": "duration_report.xlsx",
                }.get(sec, "consolidated_report.xlsx")
                _file_resp(self, export_reports_excel_bytes(sec, lob, q), XLSX_MIME, fname)
            elif path == "/export_excel/expenditures":
                _file_resp(self, export_expenditures_excel_bytes(), XLSX_MIME, "expenditures_report.xlsx")
            elif path == "/export_excel/financials":
                qs       = parse_qs(urlparse(self.path).query)
                lob      = qs.get("lob", [""])[0]
                ebf_mult = qs.get("ebf_mult", ["1.35"])[0]
                pm_sel   = qs.get("pm", [""])[0]
                _file_resp(self, export_financials_excel_bytes(lob, ebf_mult, pm_sel), XLSX_MIME, "financials_report.xlsx")
            elif path == "/export_excel/variations":
                _file_resp(self, export_variations_excel_bytes(), XLSX_MIME, "variations_report.xlsx")
            elif path == "/export_excel/kpi":
                _file_resp(self, export_kpi_excel_bytes(), XLSX_MIME, "kpi_report.xlsx")
            elif path == "/export_excel/eac_status":
                qs = parse_qs(urlparse(self.path).query)
                _file_resp(self, export_eac_status_excel_bytes(
                    pm=qs.get("pm", [""])[0],
                    status=qs.get("status", [""])[0],
                    q=qs.get("q", [""])[0],
                ), XLSX_MIME, "eac_status_report.xlsx")
            elif path == "/export_excel/schedule":
                qs = parse_qs(urlparse(self.path).query)
                _file_resp(self, export_schedule_excel_bytes(
                    lob=qs.get("lob", [""])[0],
                    pm=qs.get("pm", [""])[0],
                    status=qs.get("status", [""])[0],
                    q=qs.get("q", [""])[0],
                ), XLSX_MIME, "schedule_performance_report.xlsx")
            elif path == "/export_excel/billing":
                qs  = parse_qs(urlparse(self.path).query)
                pid = qs.get("project_id", [""])[0].strip()
                lob = qs.get("lob", [""])[0].strip()
                pf  = qs.get("pending_filter", ["all"])[0]
                own_pm = (getattr(_req_ctx, "pm_name", "") or "").strip()
                if role == "pm" and own_pm:
                    # Scoped PM: only ever export their own, already-selected project.
                    # Never fall back to an unfiltered/lob-wide export.
                    if pid.isdigit() and pm_owns_project(own_pm, int(pid)):
                        lob = ""
                    else:
                        pid = "-1"   # no project can match — yields an empty export
                        lob = ""
                _file_resp(self, export_billing_excel_bytes(
                    project_id=pid or None, lob=lob or None,
                    pending_only=(pf == "pending")
                ), XLSX_MIME, "billing_collection_report.xlsx")
            elif path == "/get_parent_data":
                import json as _json
                qs      = parse_qs(urlparse(self.path).query)
                pname   = qs.get("parent_name", [""])[0]
                excl_id = int(qs.get("exclude_id", ["0"])[0])
                data    = fetch_parent_with_children_totals(pname, excl_id) if pname else {}
                payload = _json.dumps(data).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type",   "application/json")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            elif path == "/chat_sessions":
                import json as _json
                sessions = fetch_chat_sessions(user)
                payload = _json.dumps({"sessions": sessions}, default=str).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type",   "application/json")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            elif path == "/chat_history":
                import json as _json
                qs  = parse_qs(urlparse(self.path).query)
                sid = qs.get("session_id", [""])[0].strip()
                msgs = fetch_chat_messages(sid) if sid else []
                payload = _json.dumps({"messages": msgs}, default=str).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type",   "application/json")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            elif path == "/get_parents_by_lob":
                import json as _json
                qs  = parse_qs(urlparse(self.path).query)
                lob = qs.get("lob", [""])[0].strip()
                try:
                    with ENGINE.connect() as conn:
                        if lob:
                            rows = conn.execute(text("""
                                SELECT id, COALESCE(co_no,''), project_name FROM projects
                                WHERE project_type = 'parent'
                                  AND (hier_business_line = :lob OR business_line = :lob)
                                ORDER BY project_name
                            """), {"lob": lob}).fetchall()
                        else:
                            rows = conn.execute(text("""
                                SELECT id, COALESCE(co_no,''), project_name FROM projects
                                WHERE project_type = 'parent'
                                ORDER BY project_name
                            """)).fetchall()
                    # co_no is the authoritative key (unique per project); label is what's shown/typed
                    names = [
                        {"co_no": r[1], "label": (f"{r[1]}-{r[2]}" if r[1] else r[2]), "name": r[2]}
                        for r in rows if r[2]
                    ]
                except Exception:
                    names = []
                payload = _json.dumps(names).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type",   "application/json")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            elif path == "/wipe_data":
                wipe_all_transactional_data()
                _html_resp(self, project_data_entry(
                    alert_html='<div class="alert ok">✅ All previous project, expenditure, and LOB data has been wiped from memory. You can now upload your Excel file.</div>',
                    username=user))
                return
            
            else:
              _html_resp(self,
                    _html_head("404") + "<body>" + _sidebar() + _topbar(username=user) +
                    f'<div class="main"><div class="canvas" style="padding-top:60px">'
                    f'<p style="font-size:20px;color:{C["error"]}">404 — Not found: <code>{path}</code></p>'
                    f'<p><a href="/">← Dashboard</a></p></div></div></body></html>', code=404)

        except Exception:
            _html_resp(self,
                f'<html><body><pre style="padding:32px;color:red">{traceback.format_exc()}</pre></body></html>',
                code=500)

    def do_POST(self):
        path = urlparse(self.path).path
        try:
            if path == "/login":
                fields   = _parse_body(self)
                username = fields.get("username","").strip()
                password = fields.get("password","")
                next_url = fields.get("next","/dashboard")
                user_obj = verify_user(username, password)
                if not user_obj:
                    _html_resp(self, login_page(
                        alert_html='<div class="alert err">⚠ Incorrect username or password.</div>',
                        next_path=next_url))
                    return
                token = create_session(user_obj)
                self.send_response(302)
                self.send_header("Set-Cookie",
                    f"{SESSION_COOKIE}={token}; Path=/; HttpOnly; SameSite=Lax")
                self.send_header("Location", next_url or "/dashboard")
                self.end_headers(); return

            if path == "/reset_password":
                fields   = _parse_body(self)
                username = fields.get("username","").strip()
                new_pw   = fields.get("new_password","")
                conf_pw  = fields.get("confirm_password","")

                if not username or not new_pw:
                    _html_resp(self, reset_password_page(
                        alert_html='<div class="alert-err">⚠ Please enter your username and a new password.</div>',
                        prefill_username=username))
                    return
                if new_pw != conf_pw:
                    _html_resp(self, reset_password_page(
                        alert_html='<div class="alert-err">⚠ Passwords do not match. Try again.</div>',
                        prefill_username=username))
                    return
                if len(new_pw) < 6:
                    _html_resp(self, reset_password_page(
                        alert_html='<div class="alert-err">⚠ Password must be at least 6 characters.</div>',
                        prefill_username=username))
                    return

                ok = reset_user_password(username, new_pw)
                if not ok:
                    _html_resp(self, reset_password_page(
                        alert_html='<div class="alert-err">⚠ No account found with that username.</div>',
                        prefill_username=username))
                    return

                _html_resp(self, login_page(
                    alert_html='<div class="alert-ok">✅ Password updated successfully. Please sign in with your new password.</div>'))
                return

            user = get_session_user(self)
            if not user:
                _redirect(self, "/login"); return
            role = get_session_role(self)
            _req_ctx.role     = role
            _req_ctx.is_admin = is_admin_role(role)
            _req_ctx.pm_name  = get_session_pm_name(self)
            _req_ctx.active_only = get_session_active_filter(self)

            if role == "director" and path not in DIRECTOR_ALLOWED_POST_PATHS:
                _redirect(self, "/variations"); return

            if role == "accounts" and path not in ACCOUNTS_ALLOWED_POST_PATHS:
                _redirect(self, "/payment_approvals"); return

            if role == "procurement" and path not in PROCUREMENT_ALLOWED_POST_PATHS:
                _redirect(self, "/payment_approvals"); return

            if role in ("pm", "head") and path in PM_BLOCKED_PATHS:
                _redirect(self, "/dashboard"); return

            # ── Multipart routes (read body themselves — must come before _parse_body) ──

            if path == "/upload":
                content_type = self.headers.get("Content-Type", "")
                if "multipart/form-data" in content_type:
                    try:
                        import email, email.policy, tempfile, os

                        length   = int(self.headers.get("Content-Length", 0))
                        raw_body = self.rfile.read(length)

                        full_msg_bytes = (
                            f"Content-Type: {content_type}\r\n\r\n".encode() + raw_body
                        )
                        msg = email.message_from_bytes(
                            full_msg_bytes, policy=email.policy.compat32
                        )

                        file_data = None
                        filename  = None

                        for part in msg.walk():
                            cd = part.get("Content-Disposition", "")
                            if 'name="project_file"' in cd or "name='project_file'" in cd:
                                for segment in cd.split(";"):
                                    segment = segment.strip()
                                    if segment.startswith("filename="):
                                        filename = segment.split("=", 1)[1].strip().strip('"\'')
                                file_data = part.get_payload(decode=True)
                                break

                        if file_data and filename:
                            suffix   = ".xlsx" if filename.lower().endswith(".xlsx") else ".csv"
                            tmp_path = None
                            try:
                                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                                    tmp.write(file_data)
                                    tmp_path = tmp.name
                                alert = _process_upload(tmp_path, suffix)
                            finally:
                                if tmp_path and os.path.exists(tmp_path):
                                    try:    os.unlink(tmp_path)
                                    except OSError: pass
                            _html_resp(self, project_data_entry(alert_html=alert, username=user))
                            return

                    except Exception as exc:
                        _html_resp(self, project_data_entry(
                            alert_html=f'<div class="alert err">Upload error: {exc}<br>'
                                      f'<pre style="font-size:11px">{traceback.format_exc()}</pre></div>',
                            username=user))
                        return
                _redirect(self, "/data_entry")
                return
            elif path == "/check_folder_access":
                import os, json

                try:
                    length = int(self.headers.get("Content-Length", 0))
                    raw_body = self.rfile.read(length) if length > 0 else b"{}"
                    data = json.loads(raw_body.decode("utf-8") or "{}")
                except Exception:
                    data = {}

                folder_path = (data.get("folder_path") or "").strip()

                if not folder_path:
                    result = {"accessible": False, "message": "No path provided."}
                    status_code = 400
                else:
                    try:
                        if os.path.isdir(folder_path) and os.access(folder_path, os.R_OK):
                            result = {"accessible": True}
                        else:
                            result = {"accessible": False, "message": "Path not found or insufficient permissions."}
                    except Exception:
                        result = {"accessible": False, "message": "Error while checking this path."}
                    status_code = 200

                body = json.dumps(result).encode("utf-8")
                self.send_response(status_code)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
                return

            elif path == "/wipe_data":
                wipe_all_transactional_data()
                _html_resp(self, project_data_entry(
                    alert_html='<div class="alert ok">✅ All previous project, expenditure, and LOB data has been wiped from memory. You can now upload your Excel file.</div>',
                    username=user))
                return
            
            elif path == "/upload_expenditure":
                content_type = self.headers.get("Content-Type", "")
                if "multipart/form-data" in content_type:
                    try:
                        import email, email.policy, tempfile, os

                        length   = int(self.headers.get("Content-Length", 0))
                        raw_body = self.rfile.read(length)

                        full_msg_bytes = (
                            f"Content-Type: {content_type}\r\n\r\n".encode() + raw_body
                        )
                        msg = email.message_from_bytes(
                            full_msg_bytes, policy=email.policy.compat32
                        )

                        file_data = None
                        filename  = None

                        for part in msg.walk():
                            cd = part.get("Content-Disposition", "")
                            if 'name="expenditure_file"' in cd or "name='expenditure_file'" in cd:
                                for segment in cd.split(";"):
                                    segment = segment.strip()
                                    if segment.startswith("filename="):
                                        filename = segment.split("=", 1)[1].strip().strip('"\'')
                                file_data = part.get_payload(decode=True)
                                break

                        if file_data and filename:
                            tmp_path = None
                            try:
                                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                                    tmp.write(file_data)
                                    tmp_path = tmp.name
                                alert = _process_expenditure_upload(tmp_path)
                            finally:
                                if tmp_path and os.path.exists(tmp_path):
                                    try:    os.unlink(tmp_path)
                                    except OSError: pass
                            _html_resp(self, expenditures(alert_html=alert, username=user))
                            return

                    except Exception as exc:
                        _html_resp(self, expenditures(
                            alert_html=f'<div class="alert err">Upload error: {exc}<br>'
                                      f'<pre style="font-size:11px">{traceback.format_exc()}</pre></div>',
                            username=user))
                        return
                _redirect(self, "/expenditures")
                return

            elif path == "/save_booking":
                content_type = self.headers.get("Content-Type", "")
                if "multipart/form-data" in content_type:
                    try:
                        b_fields, b_files = _parse_multipart_full(self)
                        ok, msg, booking_uid = save_payment_booking(
                            b_fields, b_files, user, role, getattr(_req_ctx, "pm_name", ""))
                        alert = (f'<div class="alert ok">✅ {msg}</div>' if ok
                                 else f'<div class="alert err">⚠ {msg}</div>')
                        edit_bid_retry = (b_fields.get("booking_id") or "").strip()
                        retry_edit_booking = fetch_booking_detail(int(edit_bid_retry)) if (not ok and edit_bid_retry.isdigit()) else None
                        _html_resp(self, weekly_payment_booking(
                            alert_html=alert, username=user, role=role,
                            pm_name=getattr(_req_ctx, "pm_name", ""),
                            sel_project_id=("" if ok else b_fields.get("project_id", "")),
                            is_indirect=(b_fields.get("is_indirect") == "1"),
                            edit_booking=retry_edit_booking,
                        ))
                        return
                    except Exception as exc:
                        _html_resp(self, weekly_payment_booking(
                            alert_html=f'<div class="alert err">Upload error: {exc}</div>',
                            username=user, role=role, pm_name=getattr(_req_ctx, "pm_name", "")))
                        return
                _redirect(self, "/weekly_booking")
                return

            elif path in ("/delete_chat_session", "/clear_chat_history"):
                import json as _json
                try:
                    content_type = self.headers.get("Content-Type", "")
                    if "application/json" in content_type:
                        length = int(self.headers.get("Content-Length", 0))
                        raw_body = self.rfile.read(length) if length > 0 else b"{}"
                        data = _json.loads(raw_body.decode("utf-8") or "{}")
                    else:
                        data = _parse_body(self)
                    if path == "/delete_chat_session":
                        sid = str(data.get("session_id", "")).strip()
                        ok = delete_chat_session(sid, user) if sid else False
                    else:
                        ok = clear_all_chat_history(user)
                    res = {"ok": ok}
                except Exception as exc:
                    res = {"ok": False, "error": str(exc)}
                payload = _json.dumps(res).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return

            # ── All other routes: parse form body first ──
            fields = _parse_body(self)

            if path == "/procurement_decide":
                bid = int(fields.get("booking_id", 0) or 0)
                if bid:
                    procurement_decide(bid, fields.get("decision", "disapprove"), fields.get("note", ""), user)
                _redirect(self, f"/payment_approvals?id={bid}")

            elif path == "/accounts_decide":
                bid = int(fields.get("booking_id", 0) or 0)
                if bid:
                    accounts_decide(bid, fields.get("decision", "disapprove"), fields.get("note", ""), user)
                _redirect(self, f"/payment_approvals?id={bid}")

            elif path == "/director_decide":
                bid = int(fields.get("booking_id", 0) or 0)
                if bid:
                    director_decide(bid, fields.get("decision", "disapprove"), fields.get("note", ""), user)
                _redirect(self, f"/payment_approvals?id={bid}")

            elif path == "/create_user":
                role = get_session_role(self)
                if not is_admin_role(role):
                    _redirect(self, "/dashboard")
                else:
                    ok, msg = create_user(
                        fields.get("username", ""),
                        fields.get("password", ""),
                        fields.get("display_name", ""),
                        fields.get("role", "user"),
                        fields.get("pm_name", ""),
                    )
                    alert = (f'<div class="alert ok">✅ {msg}</div>' if ok
                            else f'<div class="alert err">⚠ {msg}</div>')
                    _html_resp(self, user_management(alert_html=alert, username=user))

            elif path == "/delete_user":
                role = get_session_role(self)
                if not is_admin_role(role):
                    _redirect(self, "/dashboard")
                else:
                    uid = int(fields.get("user_id", 0))
                    if uid:
                        delete_user_account(uid)
                    _redirect(self, "/user_management")

            elif path == "/update_user_role":
                role = get_session_role(self)
                if not is_admin_role(role):
                    _redirect(self, "/dashboard")
                else:
                    uid = int(fields.get("user_id", 0))
                    new_role = fields.get("role", "user")
                    if uid:
                        ok = update_user_role(uid, new_role)
                        if ok:
                            alert = f'<div class="alert ok">✅ Role updated to <b>{new_role.upper()}</b> successfully.</div>'
                        else:
                            alert = '<div class="alert err">⚠ Failed to update role.</div>'
                    else:
                        alert = '<div class="alert err">⚠ Invalid user.</div>'
                    _html_resp(self, user_management(alert_html=alert, username=user))
            elif path == "/generate_pm_logins":
                role = get_session_role(self)
                if not is_admin_role(role):
                    _redirect(self, "/dashboard")
                else:
                    count, created = generate_pm_logins()
                    if count:
                        detail = "<br>".join(created)
                        alert = f'<div class="alert ok">✅ Created {count} PM login(s):<br>{detail}</div>'
                    else:
                        alert = '<div class="alert ok">Every Project Manager already has a login.</div>'
                    _html_resp(self, user_management(alert_html=alert, username=user))

            elif path == "/reset_user_password":
                role = get_session_role(self)
                if not is_admin_role(role):
                    _redirect(self, "/dashboard")
                else:
                    target_user = fields.get("username", "")
                    new_pw = fields.get("new_password", "")
                    if not target_user or not new_pw:
                        alert = '<div class="alert err">⚠ Username and new password are required.</div>'
                    elif len(new_pw) < 6:
                        alert = '<div class="alert err">⚠ Password must be at least 6 characters.</div>'
                    else:
                        ok = reset_user_password(target_user, new_pw)
                        if ok:
                            alert = f'<div class="alert ok">✅ Password for <b>{target_user}</b> has been reset.</div>'
                        else:
                            alert = f'<div class="alert err">⚠ Could not reset password for "{target_user}".</div>'
                    _html_resp(self, user_management(alert_html=alert, username=user))

            elif path == "/submit":
                import json as _json
                pid = fields.get("id", "").strip()
                is_ajax = self.headers.get("X-Requested-With") == "XMLHttpRequest"

                # fallback: doc_project_id hidden field
                if not pid:
                    pid = fields.get("doc_project_id", "").strip()

                # fallback: URL query string ?id=...
                if not pid:
                    pid = parse_qs(urlparse(self.path).query).get("id", [""])[0].strip()

                try:
                    if pid:
                        update_project(int(pid), fields)
                        project_id = int(pid)
                    else:
                        project_id = save_project(fields)
                        if not project_id:
                            with ENGINE.connect() as conn:
                                row = conn.execute(text("SELECT MAX(id) FROM projects")).fetchone()
                                project_id = row[0] if row and row[0] else 0

                    if is_ajax:
                        payload = _json.dumps({"id": project_id, "status": "ok"}).encode("utf-8")
                        self.send_response(200)
                        self.send_header("Content-Type", "application/json")
                        self.send_header("Content-Length", str(len(payload)))
                        self.end_headers()
                        self.wfile.write(payload)
                    else:
                        _redirect(self, f"/data_entry?id={project_id}")

                except Exception as e:
                    traceback.print_exc()
                    if is_ajax:
                        err = _json.dumps({"error": str(getattr(e, "orig", None) or e)[:500]}).encode("utf-8")
                        self.send_response(500)
                        self.send_header("Content-Type", "application/json")
                        self.send_header("Content-Length", str(len(err)))
                        self.end_headers()
                        self.wfile.write(err)
                    else:
                        raise

            elif path == "/delete_project":
                role_ = (getattr(_req_ctx, "role", "") or "").lower()
                if role_ == "pm":
                    _redirect(self, "/financials"); return
                pid = int(fields.get("project_id", 0))
                if pid: delete_project(pid)
                _redirect(self, "/data_entry")

            elif path == "/link_doc":
                gdrive   = fields.get("gdrive_link","").strip()
                net      = fields.get("net_path","").strip()
                doc_name = fields.get("doc_name","").strip()
                # Only insert if there is actually something meaningful to store
                if doc_name or gdrive or net:
                    parts = []
                    if gdrive: parts.append(f"GDRIVE::{gdrive}")
                    if net:    parts.append(f"NET::{net}")
                    doc_link = "|".join(parts) if parts else ""
                    link_document({
                        "project_id":   fields.get("project_id", None),
                        "doc_name":     doc_name,
                        "doc_category": fields.get("doc_category",""),
                        "doc_link":     doc_link,
                    })
                pid_str = fields.get("project_id", "").strip()
                _redirect(self, f"/data_entry?id={pid_str}#tab4" if pid_str else "/data_entry")

            elif path == "/delete_doc":
                did = int(fields.get("doc_id", 0))
                if did: delete_document(did)
                _redirect(self, "/data_entry")

            elif path == "/expenditures_retrieve":
                co  = fields.get("co_no","")
                lob = fields.get("lob","")
                pm  = fields.get("pm","")
                _redirect(self, f"/expenditures?co_no={co}&lob={lob}&pm={pm}")

            elif path == "/billing_collection_retrieve":
                pid = fields.get("project_id","")
                pm  = fields.get("pm","")
                pf  = fields.get("pending_filter","all")
                _redirect(self, f"/billing_collection?project_id={pid}&pm={pm}&pending_filter={pf}")

            elif path == "/save_billing_entry":
                pid = fields.get("project_id","")
                pm  = fields.get("pm","")
                role   = (getattr(_req_ctx, "role", "") or "").lower()
                own_pm = (getattr(_req_ctx, "pm_name", "") or "").strip()
                if role == "pm" and own_pm:
                    pm = own_pm
                    if not (pid.isdigit() and pm_owns_project(own_pm, int(pid))):
                        _redirect(self, "/billing_collection"); return
                save_billing_entry(fields)
                _redirect(self, f"/billing_collection?project_id={pid}&pm={pm}")

            elif path == "/record_collection":
                eid = int(fields.get("entry_id", 0))
                pid = fields.get("project_id","")
                pm  = fields.get("pm","")
                role   = (getattr(_req_ctx, "role", "") or "").lower()
                own_pm = (getattr(_req_ctx, "pm_name", "") or "").strip()
                if role == "pm" and own_pm:
                    pm = own_pm
                    if not pm_owns_billing_entry(own_pm, eid):
                        _redirect(self, "/billing_collection"); return
                if eid:
                    update_billing_collection(eid, fields.get("collection_date",""), fields.get("collection_amt",0))
                _redirect(self, f"/billing_collection?project_id={pid}&pm={pm}")

            elif path == "/delete_billing_entry":
                eid = int(fields.get("entry_id", 0))
                pid = fields.get("project_id","")
                pm  = fields.get("pm","")
                role   = (getattr(_req_ctx, "role", "") or "").lower()
                own_pm = (getattr(_req_ctx, "pm_name", "") or "").strip()
                if role == "pm" and own_pm:
                    pm = own_pm
                    if not pm_owns_billing_entry(own_pm, eid):
                        _redirect(self, "/billing_collection"); return
                if eid: delete_billing_entry(eid)
                _redirect(self, f"/billing_collection?project_id={pid}&pm={pm}")

            elif path == "/save_expenditure":
                result = save_expenditure(fields)
                co  = fields.get("co_no","")
                lob = fields.get("lob","")
                if not result:
                    _html_resp(self, expenditures(
                        alert_html='<div class="alert err">⚠ Please select and retrieve a project before updating ERP data.</div>',
                        prefill=fields, username=user))
                else:
                    _redirect(self, f"/expenditures?co_no={co}&lob={lob}")

            elif path == "/save_entry":
                save_expense_entry(fields)
                co = fields.get("co_no","")
                _redirect(self, f"/expenditures?co_no={co}")

            elif path == "/delete_entry":
                eid = int(fields.get("entry_id", 0))
                co  = fields.get("co_no","")
                lob = fields.get("lob","")
                pm  = fields.get("pm","")
                if eid: delete_expense_entry(eid)
                _redirect(self, f"/expenditures?co_no={co}&lob={lob}&pm={pm}" if co else "/expenditures")

            elif path == "/financials_filter":
                _html_resp(self, financials(prefill=fields, username=user))

            elif path == "/save_ebf_lob":
                lob = fields.get("lob", "").strip()
                pm_sel = fields.get("pm", "").strip()
                save_ebf_mult(lob, fields.get("ebf_mult", DEFAULT_EBF_MULT))
                _html_resp(self, financials(prefill={"lob": lob, "pm": pm_sel}, username=user))

            elif path == "/save_variation":
                try:
                    orig = _n(fields.get("original_value"))
                    vamt = _n(fields.get("variation_amt"))
                    fields["revised_value"] = orig + vamt
                except: pass
                save_variation(fields)
                _redirect(self, "/variations")

            elif path == "/delete_variation":
                vid = int(fields.get("var_id", 0))
                if vid: delete_variation(vid)
                _redirect(self, "/variations")

            elif path == "/add_ref":
                key = fields.get("ref_key","")
                val = fields.get("ref_val","").strip()
                if key in REF_TABLE_MAP and val:
                    table, col = REF_TABLE_MAP[key]
                    add_ref_value(table, col, val)
                _redirect(self, "/data_entry")

            elif path == "/delete_ref":
                key = fields.get("ref_key","")
                val = fields.get("ref_val","").strip()
                if key in REF_TABLE_MAP and val:
                    table, col = REF_TABLE_MAP[key]
                    delete_ref_value(table, col, val)
                _redirect(self, "/data_entry")

            elif path == "/add_ref_option":
                cat = fields.get("category","")
                val = fields.get("value","").strip()
                if cat in REF_TABLE_MAP and val:
                    table, col = REF_TABLE_MAP[cat]
                    add_ref_value(table, col, val)
                _redirect(self, "/data_entry")

            elif path == "/delete_ref_option":
                cat = fields.get("category","")
                val = fields.get("value","").strip()
                if cat in REF_TABLE_MAP and val:
                    table, col = REF_TABLE_MAP[cat]
                    delete_ref_value(table, col, val)
                _redirect(self, "/data_entry")

            else:
                _redirect(self, "/")

        except Exception:
            _html_resp(self,
                f'<html><body><pre style="padding:32px;color:red">{traceback.format_exc()}</pre></body></html>',
                code=500)

# ═══════════════════════════════════════════════════════════════
# SERVER BOOT — LAN ONLY
# ═══════════════════════════════════════════════════════════════
class ThreadedServer(socketserver.ThreadingTCPServer):
    allow_reuse_address = True

def get_lan_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def run_server(port=9880):

    try:
        init_db()
    except Exception:
        traceback.print_exc()
        return

    ip = get_lan_ip()
    lan_url = f"http://{ip}:{port}"

    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(base_dir)
        os.makedirs("exports", exist_ok=True)
    except Exception:
        pass

    print("\n" + "=" * 70)
    print("🚀 MONARCH SURVEYORS PMS v4.1")
    print("=" * 70)
    print(f"🌐 Login URL : {lan_url}/login")
    print(f"📊 Dashboard : {lan_url}/dashboard")
    print(f"🗄️ Database  : {DB_USER}@{DB_HOST}:{DB_PORT}/{DB_NAME}")
    print(f"📁 Exports   : {os.path.abspath('exports')}")
    print("=" * 70)
    print("Access from any PC on the same network")
    print("Press CTRL+C to stop server")
    print("=" * 70 + "\n")

    try:
        with ThreadedServer(("0.0.0.0", port), MonarchHandler) as httpd:

            try:
                webbrowser.open(lan_url + "/login")
            except Exception:
                pass

            httpd.serve_forever()

    except KeyboardInterrupt:
        print("\n🛑 Server stopped.")

    except Exception:
        traceback.print_exc()


if __name__ == "__main__":
    run_server(9880)