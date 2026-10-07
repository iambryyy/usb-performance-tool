"""Browser helper shared by PDF export and the packaged-app self-test."""


def launch_browser(p, headless):
    """
    Launch a Chromium-based browser without needing any separate download.
    Microsoft Edge ships with every Windows 10/11 install (and Playwright drives it
    like Chromium, including PDF export); Chrome is tried second, and Playwright's own
    downloaded Chromium only as a last resort. This is what makes PDF export work on a
    fresh machine with nothing installed.
    """
    last_error = None
    for channel in ("msedge", "chrome", None):
        try:
            if channel:
                return p.chromium.launch(channel=channel, headless=headless)
            return p.chromium.launch(headless=headless)
        except Exception as error:
            last_error = error
    raise last_error
