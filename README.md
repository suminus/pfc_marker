# pfc_marker

http://suminus.github.io/pfc_marker

pfc_marker imports edls (cmx3600),csv, dvs clipster-timelines as clip-frame-markers into a existing pfclean-timeline.

```
edl: source- or record-tc only, dissolves are filterd out, without comments (as of now)

clipster: timeline-markers including names and comments

csv-text: one marker-event per line, timecode separated by by the first comma
```
timecode-math done with: https://github.com/eoyilmaz/timecode

drop/non-drop handling is not implemented, yet.

written in python3.4.3 and pyqt5, windows x64


