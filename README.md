# Air Swipe

Change slides by swiping your hand in the air in front of your laptop webcam. No touching needed.

- Swipe **left** for the next slide
- Swipe **right** for the previous slide

It sends Left/Right arrow keys, so it works with PowerPoint, Google Slides, PDF viewers, image viewers and anything else that uses the arrow keys. It also works with an external monitor: the laptop webcam watches your hand, and the slides change wherever they're displayed.

## How it works

The script tracks your hand with [MediaPipe Hands](https://developers.google.com/mediapipe). When your open palm moves far enough sideways in a short time, it counts as a swipe and sends the key with `pyautogui`. A short cooldown stops one swipe from counting twice. Only an open hand counts, so normal hand movement while you talk shouldn't change slides.

## Setup

Requires Python 3.9 to 3.12 (MediaPipe doesn't support newer versions yet).

```bash
git clone https://github.com/<your-username>/air_swipe.git
cd air_swipe
python -m venv .venv
.venv\Scripts\activate        # Windows  (macOS/Linux: source .venv/bin/activate)
pip install -r requirements.txt
```

## Usage

```bash
python air_swipe.py                 # send arrow keys to the focused window
python air_swipe.py --ppt           # control a PowerPoint slideshow directly (Windows)
python air_swipe.py --no-preview    # run without the camera window
python air_swipe.py --camera 1      # use a different camera
```

1. Start the script. Hold up an open palm about 50 to 80 cm from the webcam until the camera window shows "open palm".
2. Open your presentation and start the slideshow (F5 in PowerPoint).
3. Swipe.

Press `q` in the camera window to quit, or Ctrl+C in the terminal when using `--no-preview`.

**Tip for presenting:** `python air_swipe.py --ppt --no-preview` keeps the camera window out of the way. It also changes slides even if PowerPoint isn't the active window, which helps with dual screens and Presenter View.

## Tuning

Edit the constants at the top of `air_swipe.py`:

| Setting | Default | Meaning |
|---|---|---|
| `SWIPE_THRESHOLD` | `0.20` | How far the hand must move, as a fraction of the frame width |
| `SWIPE_WINDOW` | `0.4` | Seconds within which that movement must happen |
| `COOLDOWN` | `0.8` | Seconds to ignore movement after a swipe |
| `REQUIRE_OPEN_PALM` | `True` | Only count swipes made with an open hand |
| `INVERT_DIRECTION` | `False` | Set to `True` to swipe right for the next slide |

- **Swipes are missed:** lower `SWIPE_THRESHOLD` or raise `SWIPE_WINDOW`.
- **Slides change by accident:** raise `SWIPE_THRESHOLD`.
- **One swipe skips two slides:** raise `COOLDOWN`.

## Notes

- If the camera window is the active window, the arrow keys go to it instead of your slides. Click back on the slideshow, or use `--ppt` / `--no-preview`.
- `pyautogui`'s fail-safe is on: moving the mouse into a screen corner stops the script.
