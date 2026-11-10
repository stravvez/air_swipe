"""
Air Swipe - control slides by swiping your hand in front of the webcam.

  Swipe hand LEFT  -> next slide     (Right arrow)
  Swipe hand RIGHT -> previous slide (Left arrow)

Usage:
  python air_swipe.py              # sends arrow keys to the focused window (any app)
  python air_swipe.py --ppt        # drives a running PowerPoint slideshow directly,
                                   # even if PowerPoint isn't the focused window
  python air_swipe.py --no-preview # no camera window
  python air_swipe.py --camera 1   # use a different camera

Press 'q' in the preview window (or Ctrl+C in the terminal) to quit.
Works with dual monitors: the webcam is only the input; keys go to the focused
window on whichever screen it is.
"""

import argparse
import time
from collections import deque

import cv2
import mediapipe as mp
import pyautogui

# ---------------- Tuning ----------------
SWIPE_THRESHOLD = 0.20   # fraction of frame width the hand must travel
SWIPE_WINDOW = 0.4       # seconds within which that travel must happen
COOLDOWN = 0.3           # seconds to ignore movement after a swipe
REQUIRE_OPEN_PALM = True # only count swipes with an open hand (fewer false triggers)
INVERT_DIRECTION = False # True = swipe right for next slide
# ----------------------------------------

pyautogui.FAILSAFE = True  # slam mouse into a screen corner to abort
pyautogui.PAUSE = 0

mp_hands = mp.solutions.hands
mp_draw = mp.solutions.drawing_utils

FINGERS = [(8, 6), (12, 10), (16, 14), (20, 18)]  # (tip, pip) for index..pinky


def is_open_palm(lm):
    """Finger counts as extended if its tip is farther from the wrist than its PIP joint.
    Distance-based so it works even when the hand is tilted mid-swipe."""
    wrist = lm[0]

    def d2(p):
        return (p.x - wrist.x) ** 2 + (p.y - wrist.y) ** 2

    extended = sum(1 for tip, pip in FINGERS if d2(lm[tip]) > d2(lm[pip]) * 1.15)
    return extended >= 3


class PowerPointController:
    """Talks to PowerPoint via COM so the slideshow moves without needing focus."""

    def __init__(self):
        import win32com.client  # from pywin32
        self._client = win32com.client

    def _view(self):
        try:
            app = self._client.GetActiveObject("PowerPoint.Application")
            if app.SlideShowWindows.Count > 0:
                return app.SlideShowWindows(1).View
        except Exception:
            pass
        return None

    def go(self, forward):
        view = self._view()
        if view is None:
            return False
        view.Next() if forward else view.Previous()
        return True


def send(forward, ppt):
    if ppt is not None and ppt.go(forward):
        return "PowerPoint"
    pyautogui.press("right" if forward else "left")
    return "keys"


def main():
    ap = argparse.ArgumentParser(description="Swipe your hand to change slides.")
    ap.add_argument("--camera", type=int, default=0)
    ap.add_argument("--no-preview", action="store_true")
    ap.add_argument("--ppt", action="store_true",
                    help="control PowerPoint slideshow directly (falls back to keys)")
    args = ap.parse_args()

    ppt = None
    if args.ppt:
        try:
            ppt = PowerPointController()
        except ImportError:
            print("pywin32 not installed; falling back to arrow keys. (pip install pywin32)")

    cap = cv2.VideoCapture(args.camera, cv2.CAP_DSHOW)
    if not cap.isOpened():
        cap = cv2.VideoCapture(args.camera)
    if not cap.isOpened():
        raise SystemExit(f"Could not open camera {args.camera}")
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)

    history = deque()  # (timestamp, x)
    last_swipe = 0.0
    last_action = ""
    last_action_time = 0.0

    print("Air Swipe running. Swipe LEFT = next, RIGHT = previous. 'q' / Ctrl+C to quit.")

    with mp_hands.Hands(max_num_hands=1,
                        model_complexity=0,
                        min_detection_confidence=0.7,
                        min_tracking_confidence=0.6) as hands:
        try:
            while True:
                ok, frame = cap.read()
                if not ok:
                    print("Camera frame not received; retrying...")
                    time.sleep(0.1)
                    continue

                frame = cv2.flip(frame, 1)  # mirror so movement matches your view
                now = time.time()
                rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                result = hands.process(rgb)

                status = "no hand"
                if result.multi_hand_landmarks:
                    lm_obj = result.multi_hand_landmarks[0]
                    lm = lm_obj.landmark
                    open_palm = is_open_palm(lm)
                    status = "open palm" if open_palm else "closed hand"

                    if open_palm or not REQUIRE_OPEN_PALM:
                        history.append((now, lm[9].x))  # middle-finger base = palm center
                    else:
                        history.clear()

                    while history and now - history[0][0] > SWIPE_WINDOW:
                        history.popleft()

                    if now - last_swipe > COOLDOWN and len(history) >= 2:
                        dx = history[-1][1] - history[0][1]
                        if abs(dx) > SWIPE_THRESHOLD:
                            moved_left = dx < 0
                            forward = moved_left != INVERT_DIRECTION
                            via = send(forward, ppt)
                            last_action = "NEXT >>" if forward else "<< PREV"
                            last_action_time = now
                            last_swipe = now
                            history.clear()
                            print(f"{time.strftime('%H:%M:%S')}  {last_action}  ({via})")

                    if not args.no_preview:
                        mp_draw.draw_landmarks(frame, lm_obj, mp_hands.HAND_CONNECTIONS)
                else:
                    history.clear()

                if not args.no_preview:
                    cooling = now - last_swipe < COOLDOWN
                    cv2.putText(frame, f"{status}{'  (cooldown)' if cooling else ''}",
                                (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 255), 2)
                    if now - last_action_time < 1.0:
                        cv2.putText(frame, last_action, (10, 80),
                                    cv2.FONT_HERSHEY_SIMPLEX, 1.6, (0, 255, 0), 4)
                    cv2.imshow("Air Swipe (q to quit)", frame)
                    if cv2.waitKey(1) & 0xFF == ord("q"):
                        break
        except KeyboardInterrupt:
            pass
        finally:
            cap.release()
            cv2.destroyAllWindows()


if __name__ == "__main__":
    main()
