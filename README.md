# pfc_marker

http://suminus.github.io/pfc_marker

pfc_marker imports edls (cmx3600),csv, dvs clipster-timelines as clip-frame-markers into a existing pfclean-timeline.

```
edl: source- or record-tc only, dissolves are filterd out, tapename mapped as notes

clipster: timeline-markers including names and comments

csv-text: one marker-event per line, timecode separated by by the first comma

avid mediacomposer: mediacomposer markers exported as txt, last number shows markerrange
```
timecode-math done with: https://github.com/eoyilmaz/timecode

drop/non-drop handling is not implemented, yet.

written in python3.4.3 and pyqt5, windows 7/8/10 x64


