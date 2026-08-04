"""GUI regression tests for the layout customizer drag/resize interaction.

These need a real Tk display, so they are skipped automatically when Tk cannot
initialise (e.g. CI without an X server). Run locally under Xvfb:

    PYTHONPATH=. xvfb-run -a pytest tests/test_ui_layout.py -q
"""
from __future__ import annotations

import pytest

tk = pytest.importorskip("tkinter")


class _Ev:
    def __init__(self, x: float, y: float) -> None:
        self.x = x
        self.y = y


@pytest.fixture()
def dialog():
    import ui.app as app_mod

    try:
        app = app_mod.App()
    except tk.TclError as exc:  # no display available
        pytest.skip(f"no Tk display: {exc}")
    app.update()
    dlg = app_mod.LayoutDialog(app)
    dlg.update()
    yield dlg
    dlg.destroy()
    app.destroy()


def test_resize_tracks_full_gesture_not_one_pixel(dialog):
    """A single press→drag must move the corner the whole way, not ~1px.

    Regression: motion/release used to be bound per canvas item, so the first
    redraw destroyed the pressed handle and the drag stuck after one event."""
    box0 = list(dialog._boxes["body"])
    sx, sy = box0[2], box0[3]  # SE corner
    dialog._press(_Ev(sx, sy))
    assert dialog._drag is not None and dialog._drag[1] == "resize"
    dialog._motion(_Ev(sx - 90, sy - 50))
    box1 = list(dialog._boxes["body"])
    dialog._release(_Ev(sx - 90, sy - 50))
    # the SE corner should have moved by ~the gesture, well beyond 1px
    assert box0[2] - box1[2] > 40
    assert box0[3] - box1[3] > 20


def test_move_tracks_full_gesture(dialog):
    box0 = list(dialog._boxes["title"])
    cx, cy = (box0[0] + box0[2]) / 2, (box0[1] + box0[3]) / 2
    dialog._press(_Ev(cx, cy))
    dialog._motion(_Ev(cx, cy + 40))
    box1 = list(dialog._boxes["title"])
    dialog._release(_Ev(cx, cy + 40))
    assert box1[1] - box0[1] > 20  # moved down by ~the gesture
    # size preserved on a move
    assert round(box1[2] - box1[0]) == round(box0[2] - box0[0])


def test_min_box_size_is_enforced_on_resize(dialog):
    box0 = list(dialog._boxes["section"])
    sx, sy = box0[2], box0[3]
    dialog._press(_Ev(sx, sy))
    dialog._motion(_Ev(box0[0] - 500, box0[1] - 500))  # try to collapse
    box1 = list(dialog._boxes["section"])
    from ui.app import _MIN_BOX_PX

    assert box1[2] - box1[0] >= _MIN_BOX_PX
    assert box1[3] - box1[1] >= _MIN_BOX_PX
