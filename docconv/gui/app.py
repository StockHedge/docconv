"""그래픽 인터페이스 (tkinter).

tkinter를 고른 이유는 **파이썬 표준 라이브러리라 추가 설치가 필요 없고**
Windows/macOS/Linux에서 그대로 돌기 때문이다. PySide6가 더 예쁘지만 100MB
남짓을 더 받아야 하고, 이 프로그램의 목적(로컬에서 조용히 도는 변환기)에는
과하다.

끌어다 놓기는 `tkinterdnd2` 가 설치되어 있으면 활성화되고, 없으면 버튼으로
파일을 고르면 된다. 없다고 해서 기능이 막히지는 않는다.

변환은 별도 스레드에서 돌린다. 메인 스레드에서 돌리면 큰 파일을 변환하는
동안 창이 얼어붙어 "죽은 프로그램"처럼 보인다.
"""

from __future__ import annotations

import os
import platform
import queue
import sys
import threading
import traceback
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .. import formats, pipeline
from ..options import (
    ConvertOptions,
    MergedCellPolicy,
    OnConflict,
    SheetLayout,
    TableFlatten,
)

APP_TITLE = "문서 변환기"
IS_MAC = platform.system() == "Darwin"
IS_WIN = platform.system() == "Windows"


# --------------------------------------------------------------------------
# 끌어다 놓기 지원 (선택)
# --------------------------------------------------------------------------


def _make_root() -> tuple[tk.Tk, bool]:
    """DnD를 지원하면 TkinterDnD.Tk를, 아니면 평범한 Tk를 만든다."""
    try:
        from tkinterdnd2 import TkinterDnD

        return TkinterDnD.Tk(), True
    except Exception:
        return tk.Tk(), False


def _parse_drop(data: str) -> list[str]:
    """DnD 이벤트 문자열을 경로 목록으로.

    공백이 든 경로는 중괄호로 감싸여 온다: ``{C:/a b/c.hwp} D:/d.pdf``
    """
    out: list[str] = []
    buf = ""
    depth = 0
    for ch in data:
        if ch == "{":
            depth += 1
            if depth == 1:
                continue
        elif ch == "}":
            depth -= 1
            if depth == 0:
                out.append(buf)
                buf = ""
                continue
        if depth == 0 and ch == " ":
            if buf:
                out.append(buf)
                buf = ""
            continue
        buf += ch
    if buf:
        out.append(buf)
    return [p for p in out if p]


# --------------------------------------------------------------------------
# 앱
# --------------------------------------------------------------------------


class ConverterApp:
    def __init__(self) -> None:
        self.root, self.has_dnd = _make_root()
        self._setup_dpi()
        self.root.title(APP_TITLE)
        self.root.geometry("880x620")
        self.root.minsize(720, 520)

        self.files: list[Path] = []
        self.worker: Optional[threading.Thread] = None
        self.stop_flag = threading.Event()
        self.msg_q: "queue.Queue[tuple[str, object]]" = queue.Queue()

        self._build_vars()
        self._build_ui()
        self._enable_dnd()
        self.root.after(80, self._drain_queue)

    # -- 초기화 -----------------------------------------------------------

    @staticmethod
    def _setup_dpi() -> None:
        """Windows 고해상도 화면에서 글자가 뭉개지지 않게 한다."""
        if not IS_WIN:
            return
        try:
            import ctypes

            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    def _build_vars(self) -> None:
        self.v_target = tk.StringVar(value="pdf")
        self.v_outdir = tk.StringVar(value="")
        self.v_same_dir = tk.BooleanVar(value=True)
        self.v_keep_format = tk.BooleanVar(value=True)
        self.v_images = tk.BooleanVar(value=True)
        self.v_conflict = tk.StringVar(value=OnConflict.RENAME.value)
        self.v_merged = tk.StringVar(value=MergedCellPolicy.PLACEHOLDER.value)
        self.v_tables_only = tk.BooleanVar(value=False)
        self.v_single_sheet = tk.BooleanVar(value=False)
        self.v_csv_enc = tk.StringVar(value="utf-8-sig")
        self.v_status = tk.StringVar(value="파일을 추가하세요.")
        self.v_progress = tk.DoubleVar(value=0.0)

    def _build_ui(self) -> None:
        style = ttk.Style()
        if IS_WIN and "vista" in style.theme_names():
            style.theme_use("vista")
        elif IS_MAC and "aqua" in style.theme_names():
            style.theme_use("aqua")

        outer = ttk.Frame(self.root, padding=10)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=3)
        outer.rowconfigure(5, weight=1)

        # 1) 안내 --------------------------------------------------------
        hint = (
            "파일을 끌어다 놓거나 [파일 추가]를 누르세요."
            if self.has_dnd
            else "[파일 추가]로 변환할 파일을 고르세요."
            "   (tkinterdnd2를 설치하면 끌어다 놓기가 켜집니다)"
        )
        ttk.Label(outer, text=hint, foreground="#666").grid(
            row=0, column=0, sticky="w", pady=(0, 6)
        )

        # 2) 파일 목록 ---------------------------------------------------
        list_frame = ttk.Frame(outer)
        list_frame.grid(row=1, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        cols = ("name", "type", "size", "status")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings", height=10)
        for cid, text, w, anchor in (
            ("name", "파일", 360, "w"),
            ("type", "형식", 90, "center"),
            ("size", "크기", 90, "e"),
            ("status", "상태", 260, "w"),
        ):
            self.tree.heading(cid, text=text)
            self.tree.column(cid, width=w, anchor=anchor)
        self.tree.grid(row=0, column=0, sticky="nsew")

        sb = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.tag_configure("ok", foreground="#0a7d28")
        self.tree.tag_configure("fail", foreground="#c02020")
        self.tree.tag_configure("warn", foreground="#a06000")

        # 3) 파일 버튼 ---------------------------------------------------
        btns = ttk.Frame(outer)
        btns.grid(row=2, column=0, sticky="ew", pady=(6, 10))
        ttk.Button(btns, text="파일 추가", command=self.add_files).pack(side="left")
        ttk.Button(btns, text="폴더 추가", command=self.add_folder).pack(
            side="left", padx=4
        )
        ttk.Button(btns, text="선택 제거", command=self.remove_selected).pack(
            side="left"
        )
        ttk.Button(btns, text="모두 비우기", command=self.clear_files).pack(
            side="left", padx=4
        )
        ttk.Button(btns, text="결과 폴더 열기", command=self.open_output).pack(
            side="right"
        )

        # 4) 설정 --------------------------------------------------------
        opt = ttk.LabelFrame(outer, text="변환 설정", padding=10)
        opt.grid(row=3, column=0, sticky="ew")
        for c in range(6):
            opt.columnconfigure(c, weight=(1 if c in (1, 4) else 0))

        ttk.Label(opt, text="변환 형식").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.cb_target = ttk.Combobox(
            opt,
            textvariable=self.v_target,
            state="readonly",
            width=12,
            values=[f for f in formats.writable_formats()],
        )
        self.cb_target.grid(row=0, column=1, sticky="w")

        ttk.Label(opt, text="저장 위치").grid(row=0, column=2, sticky="w", padx=(16, 6))
        self.e_outdir = ttk.Entry(opt, textvariable=self.v_outdir)
        self.e_outdir.grid(row=0, column=3, columnspan=2, sticky="ew")
        ttk.Button(opt, text="찾아보기", command=self.pick_outdir, width=9).grid(
            row=0, column=5, padx=(6, 0)
        )
        ttk.Checkbutton(
            opt,
            text="원본과 같은 폴더에 저장",
            variable=self.v_same_dir,
            command=self._toggle_outdir,
        ).grid(row=1, column=3, columnspan=3, sticky="w", pady=(4, 0))

        ttk.Separator(opt, orient="horizontal").grid(
            row=2, column=0, columnspan=6, sticky="ew", pady=8
        )

        ttk.Checkbutton(opt, text="서식 유지", variable=self.v_keep_format).grid(
            row=3, column=0, columnspan=2, sticky="w"
        )
        ttk.Checkbutton(opt, text="이미지 포함", variable=self.v_images).grid(
            row=3, column=2, sticky="w", padx=(16, 0)
        )
        ttk.Checkbutton(opt, text="표만 추출", variable=self.v_tables_only).grid(
            row=3, column=3, sticky="w"
        )
        ttk.Checkbutton(
            opt, text="시트 하나로 합치기", variable=self.v_single_sheet
        ).grid(row=3, column=4, columnspan=2, sticky="w")

        ttk.Label(opt, text="이름 충돌").grid(row=4, column=0, sticky="w", pady=(6, 0))
        ttk.Combobox(
            opt,
            textvariable=self.v_conflict,
            state="readonly",
            width=12,
            values=[e.value for e in OnConflict],
        ).grid(row=4, column=1, sticky="w", pady=(6, 0))

        ttk.Label(opt, text="병합 셀").grid(
            row=4, column=2, sticky="w", padx=(16, 6), pady=(6, 0)
        )
        ttk.Combobox(
            opt,
            textvariable=self.v_merged,
            state="readonly",
            width=12,
            values=[e.value for e in MergedCellPolicy],
        ).grid(row=4, column=3, sticky="w", pady=(6, 0))

        ttk.Label(opt, text="CSV 인코딩").grid(
            row=4, column=4, sticky="e", padx=(16, 6), pady=(6, 0)
        )
        ttk.Combobox(
            opt,
            textvariable=self.v_csv_enc,
            state="readonly",
            width=12,
            values=["utf-8-sig", "utf-8", "cp949", "euc-kr"],
        ).grid(row=4, column=5, sticky="w", pady=(6, 0))

        # 5) 실행 --------------------------------------------------------
        run = ttk.Frame(outer)
        run.grid(row=4, column=0, sticky="ew", pady=10)
        run.columnconfigure(1, weight=1)

        self.btn_run = ttk.Button(run, text="변환 시작", command=self.start)
        self.btn_run.grid(row=0, column=0)
        self.pb = ttk.Progressbar(run, variable=self.v_progress, maximum=100.0)
        self.pb.grid(row=0, column=1, sticky="ew", padx=10)
        self.btn_stop = ttk.Button(
            run, text="중지", command=self.stop, state="disabled"
        )
        self.btn_stop.grid(row=0, column=2)

        # 6) 로그 --------------------------------------------------------
        log_frame = ttk.LabelFrame(outer, text="진행 상황", padding=6)
        log_frame.grid(row=5, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log = tk.Text(log_frame, height=7, wrap="word", state="disabled")
        self.log.grid(row=0, column=0, sticky="nsew")
        lsb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        lsb.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=lsb.set)

        ttk.Label(outer, textvariable=self.v_status, foreground="#444").grid(
            row=6, column=0, sticky="w", pady=(6, 0)
        )

        self._toggle_outdir()
        self._log(pipeline.diagnose())

    def _enable_dnd(self) -> None:
        if not self.has_dnd:
            return
        try:
            from tkinterdnd2 import DND_FILES

            self.tree.drop_target_register(DND_FILES)
            self.tree.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            self.has_dnd = False

    def _on_drop(self, event) -> None:
        paths = [Path(p) for p in _parse_drop(event.data)]
        self._add_paths(paths)

    # -- 파일 관리 ---------------------------------------------------------

    def add_files(self) -> None:
        exts = sorted(
            {e for s in formats.FORMATS.values() if s.can_read for e in s.ext}
        )
        types = [
            ("지원하는 모든 문서", " ".join("*" + e for e in exts)),
            ("한글 문서", "*.hwp *.hwpx"),
            ("Word 문서", "*.doc *.docx"),
            ("Excel 문서", "*.xls *.xlsx *.csv"),
            ("PDF", "*.pdf"),
            ("모든 파일", "*.*"),
        ]
        picked = filedialog.askopenfilenames(title="변환할 파일 선택", filetypes=types)
        self._add_paths([Path(p) for p in picked])

    def add_folder(self) -> None:
        d = filedialog.askdirectory(title="폴더 선택")
        if not d:
            return
        folder = Path(d)
        found = [
            p
            for p in sorted(folder.rglob("*"))
            if p.is_file() and formats.format_of_ext(p.suffix)
        ]
        if not found:
            messagebox.showinfo(
                APP_TITLE, "이 폴더에서 변환할 수 있는 파일을 찾지 못했습니다."
            )
            return
        self._add_paths(found)

    def _add_paths(self, paths: list[Path]) -> None:
        added = skipped = 0
        for p in paths:
            if p.is_dir():
                for f in sorted(p.rglob("*")):
                    if (
                        f.is_file()
                        and formats.format_of_ext(f.suffix)
                        and f not in self.files
                    ):
                        self.files.append(f)
                        self._insert_row(f)
                        added += 1
                continue
            if not p.is_file():
                continue
            if p in self.files:
                continue
            if not formats.format_of_ext(p.suffix):
                skipped += 1
                continue
            self.files.append(p)
            self._insert_row(p)
            added += 1

        msg = f"{added}개 추가"
        if skipped:
            msg += f", 지원하지 않는 형식 {skipped}개 제외"
        self.v_status.set(f"{msg}. 총 {len(self.files)}개")

    def _insert_row(self, p: Path) -> None:
        try:
            fmt = formats.detect(p)
            label = formats.spec(fmt).key
        except Exception:
            label = p.suffix.lstrip(".") or "?"
        try:
            size = _human_size(p.stat().st_size)
        except OSError:
            size = "-"
        self.tree.insert("", "end", iid=str(p), values=(p.name, label, size, "대기"))

    def remove_selected(self) -> None:
        for iid in self.tree.selection():
            self.tree.delete(iid)
            p = Path(iid)
            if p in self.files:
                self.files.remove(p)
        self.v_status.set(f"총 {len(self.files)}개")

    def clear_files(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self.files.clear()
        self.v_progress.set(0)
        self.v_status.set("파일을 추가하세요.")

    def open_output(self) -> None:
        target = self._out_dir()
        if target is None:
            if not self.files:
                messagebox.showinfo(APP_TITLE, "먼저 파일을 추가하세요.")
                return
            target = self.files[0].parent
        _open_folder(target)

    # -- 설정 --------------------------------------------------------------

    def _toggle_outdir(self) -> None:
        state = "disabled" if self.v_same_dir.get() else "normal"
        self.e_outdir.configure(state=state)

    def pick_outdir(self) -> None:
        d = filedialog.askdirectory(title="저장 위치 선택")
        if d:
            self.v_outdir.set(d)
            self.v_same_dir.set(False)
            self._toggle_outdir()

    def _out_dir(self) -> Optional[Path]:
        if self.v_same_dir.get():
            return None
        v = self.v_outdir.get().strip()
        return Path(v) if v else None

    def _options(self) -> ConvertOptions:
        return ConvertOptions(
            on_conflict=OnConflict(self.v_conflict.get()),
            keep_formatting=self.v_keep_format.get(),
            include_images=self.v_images.get(),
            csv_encoding=self.v_csv_enc.get(),
            merged_cells=MergedCellPolicy(self.v_merged.get()),
            table_flatten=(
                TableFlatten.TABLES_ONLY
                if self.v_tables_only.get()
                else TableFlatten.ALL
            ),
            sheet_layout=(
                SheetLayout.SINGLE
                if self.v_single_sheet.get()
                else SheetLayout.PER_SHEET
            ),
        )

    # -- 실행 --------------------------------------------------------------

    def start(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            return
        if not self.files:
            messagebox.showinfo(APP_TITLE, "변환할 파일을 먼저 추가하세요.")
            return

        target = self.v_target.get()
        out_dir = self._out_dir()
        if out_dir is not None:
            try:
                out_dir.mkdir(parents=True, exist_ok=True)
            except OSError as e:
                messagebox.showerror(APP_TITLE, f"저장 폴더를 만들 수 없습니다.\n{e}")
                return

        for p in self.files:
            if self.tree.exists(str(p)):
                self.tree.set(str(p), "status", "대기")
                self.tree.item(str(p), tags=())

        self.stop_flag.clear()
        self.v_progress.set(0)
        self.btn_run.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self._log(f"\n=== {len(self.files)}개 파일을 {target.upper()}(으)로 변환 ===")

        opts = self._options()
        files = list(self.files)
        self.worker = threading.Thread(
            target=self._run_worker, args=(files, target, out_dir, opts), daemon=True
        )
        self.worker.start()

    def _run_worker(self, files, target, out_dir, opts) -> None:
        done = 0
        ok = fail = 0
        for p in files:
            if self.stop_flag.is_set():
                self.msg_q.put(("log", "사용자 요청으로 중지했습니다."))
                break
            self.msg_q.put(("status", f"변환 중: {p.name}"))
            self.msg_q.put(("row", (str(p), "변환 중...", "")))
            try:
                r = pipeline.convert(p, target, out_dir=out_dir, opts=opts)
            except Exception as e:  # 예상 못한 오류로 전체가 멈추지 않게
                self.msg_q.put(("row", (str(p), f"오류: {e}", "fail")))
                self.msg_q.put(
                    ("log", f"[오류] {p.name}: {e}\n{traceback.format_exc(limit=3)}")
                )
                fail += 1
            else:
                if r.ok:
                    ok += 1
                    note = r.output.name if r.output else ""
                    if r.extra_outputs:
                        note += f" (+{len(r.extra_outputs)})"
                    tag = "warn" if r.warnings else "ok"
                    self.msg_q.put(("row", (str(p), f"완료 → {note}", tag)))
                    self.msg_q.put(("log", "  " + r.summary()))
                    for w in r.warnings:
                        self.msg_q.put(("log", f"    · {w}"))
                else:
                    fail += 1
                    first = (r.error or "실패").split("\n")[0]
                    self.msg_q.put(("row", (str(p), first, "fail")))
                    self.msg_q.put(("log", f"  실패  {p.name}\n    {r.error}"))
            done += 1
            self.msg_q.put(("progress", done / max(1, len(files)) * 100))
        self.msg_q.put(("done", (ok, fail)))

    def stop(self) -> None:
        self.stop_flag.set()
        self.v_status.set("중지 요청됨. 현재 파일을 마치는 중...")

    # -- 메인 스레드 갱신 ---------------------------------------------------

    def _drain_queue(self) -> None:
        """워커 스레드의 메시지를 UI에 반영한다.

        tkinter 위젯은 메인 스레드에서만 만져야 하므로 큐를 경유한다.
        """
        try:
            while True:
                kind, payload = self.msg_q.get_nowait()
                if kind == "log":
                    self._log(str(payload))
                elif kind == "status":
                    self.v_status.set(str(payload))
                elif kind == "progress":
                    self.v_progress.set(float(payload))
                elif kind == "row":
                    iid, text, tag = payload  # type: ignore[misc]
                    if self.tree.exists(iid):
                        self.tree.set(iid, "status", text)
                        self.tree.item(iid, tags=(tag,) if tag else ())
                elif kind == "done":
                    ok, fail = payload  # type: ignore[misc]
                    self.btn_run.configure(state="normal")
                    self.btn_stop.configure(state="disabled")
                    self.v_status.set(f"완료: 성공 {ok}건, 실패 {fail}건")
                    self._log(f"=== 성공 {ok}건 / 실패 {fail}건 ===")
        except queue.Empty:
            pass
        self.root.after(80, self._drain_queue)

    def _log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    # -- 진입 --------------------------------------------------------------

    def run(self) -> int:
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()
        return 0

    def _on_close(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            if not messagebox.askokcancel(
                APP_TITLE, "변환이 진행 중입니다. 종료할까요?"
            ):
                return
            self.stop_flag.set()
        self.root.destroy()


# --------------------------------------------------------------------------


def _human_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f}{unit}" if unit == "B" else f"{n:.1f}{unit}"
        n /= 1024.0
    return f"{n:.1f}GB"


def _open_folder(path: Path) -> None:
    try:
        if IS_WIN:
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif IS_MAC:
            import subprocess

            subprocess.run(["open", str(path)], check=False)
        else:
            import subprocess

            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception as e:
        messagebox.showwarning(APP_TITLE, f"폴더를 열지 못했습니다.\n{e}")


def run() -> int:
    app = ConverterApp()
    return app.run()


if __name__ == "__main__":  # pragma: no cover
    sys.exit(run())
