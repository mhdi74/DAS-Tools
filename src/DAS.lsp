(vl-load-com)

;;; ============================================================
;;;   PROGRAM NAME: DAS Tools
;;;   VERSION: 10.7.3
;;;   DEVELOPER: Seyyed Mehdi Hashemi
;;;
;;;   CHANGELOG 9.6 -> 10.0
;;;   [FIX-1] c:dasxy : oldOsm/oldClay/oldCmd were never assigned ->
;;;           (setvar 'clayer nil) crashed at the end of every run.
;;;   [FIX-2] das-core : *error* was not declared local, so it
;;;           permanently replaced the global *error* for the whole
;;;           AutoCAD session. Also localised best-idx, best-pt,
;;;           v-idx, cur-pt, loop-i, rearranged, nb, txt.
;;;           das-draw-table : localised rowsInThisCol.
;;;   [FIX-3] Dialog input is now validated/clamped. col-count = 0
;;;           caused a divide-by-zero crash; scale = 0 produced nil
;;;           text entities which then crashed the text-mask routine.
;;;   [FIX-4] das-math-utm-to-latlon : the D^5 (longitude) and D^6
;;;           (latitude) Snyder terms were subtracted instead of
;;;           added. The longitude error was ~0.5 m on the ground.
;;;   [FIX-5] Excel: vlax-get-or-create-object could hijack and then
;;;           Quit the user's own Excel session, losing unsaved work.
;;;           Now creates a private instance, wrapped in error traps
;;;           so it is always released.
;;;   [FIX-6] Excel temp2: coordinates are written as real numbers
;;;           (not rtos strings) and the SQRT length formulas in
;;;           column E are preserved instead of overwritten.
;;;   [FIX-7] Output filenames are sanitised (applicant name may
;;;           contain \ / : * ? " < > | which broke SaveAs).
;;;   [FIX-8] Styles and layers are created with entmake instead of
;;;           the -STYLE / LAYER commands. -STYLE prompts for
;;;           "Vertical?" on SHX fonts but not on TTF, so the old
;;;           code leaked one extra "N" into the next command when
;;;           creating DAS_FARSI (arial.ttf).
;;;   [NEW-1] Start-vertex rule is now selectable: NorthWest (per the
;;;           registration guideline), Northernmost (9.6 behaviour),
;;;           or AsDrawn. See notes at das-geom-start-index.
;;;   [NEW-2] The Length column is a dialog toggle. DASL still works.
;;;   [NEW-3] The duplicated 60-line draw block (preview vs final)
;;;           is now a single function, das-render-all.
;;;
;;;   CHANGELOG 10.0 -> 10.1
;;;   [FIX-9] das-math-utm-to-latlon : the meridional arc series was
;;;           missing the 3 and 5 coefficients -
;;;           a[1 - e2/4 - 3e2^2/64 - 5e2^3/256]. Latitude came out a
;;;           constant ~5.7 m too far south, which looks plausible on a
;;;           certificate and so was never noticed. Verified against
;;;           pyproj on three real parcels: 572 cm -> 0.01 cm.
;;;   [NEW-4] Corner filter, threshold *das-corner-max-angle* = 175 deg.
;;;           A vertex lying on a straight side can no longer become
;;;           vertex number 1. Points are NOT deleted - they remain in
;;;           the drawing, the table and both Excel files.
;;;           Mirrors max_angle in cadastre.py on utmdoc.ir.
;;;   [NEW-5] The NorthWest rule is the metric variant. utmdoc.ir must
;;;           run START_VERTEX_RULE=nw_metric to agree. Added a 1 cm
;;;           tie-break in favour of the western vertex.
;;;   [FIX-10] L-NEIGHBOR colour 3 (green) -> 7 (black/white).
;;;   [FIX-11] Area label is MTEXT so the square metre sign prints as a
;;;           real superscript instead of "m2".
;;;   [FIX-12] The coordinate table was placed a fixed 5 units below the
;;;           polygon, ignoring the neighbour captions drawn outside the
;;;           boundary, so on parcels with long captions the two
;;;           overlapped. Neighbours are now drawn first and the table is
;;;           positioned below the measured extents of everything drawn.
;;;   [FIX-13] Text mask called "_.-TEXTMASK", which is not a command in
;;;           any AutoCAD release - Express Tools spells it TEXTMASK,
;;;           and LT has no Express Tools at all. Replaced with native
;;;           WIPEOUT rectangles plus DRAWORDER, so the mask works on a
;;;           bare AutoCAD install.
;;;
;;;   CHANGELOG 10.7.2 -> 10.7.3
;;;   [FIX-24] English dialog: "&" in a DCL label marks the next letter as
;;;           a keyboard mnemonic and is not shown, so three group titles
;;;           read "Text  Dimension Settings". Now "and".
;;;   [NEW-18] One line after the load banner points to the support page
;;;           (https://hamibash.com/dastools). Latin only: the command line
;;;           lays text out left to right. Nothing is shown while working.
;;;
;;;   CHANGELOG 10.7.1 -> 10.7.2
;;;   [NEW-17] The command-line question "Choose Interface Language
;;;           [Farsi/English]" before every dialog is gone. The dialog opens
;;;           in Farsi in a new session (or in the language last chosen in
;;;           this session) and has a button, "English" / "فارسی", that
;;;           reopens it in the other language keeping everything typed
;;;           (das-ui-dcl-accept-apply with status 3, das-core toggles
;;;           *das-lang* and loops). Messages follow the same choice.
;;;
;;;   CHANGELOG 10.7 -> 10.7.1
;;;   [FIX-23] Area label: layer L-AREA white/black (7) instead of yellow;
;;;           preferred height 2x text (was 3x); anchored at the point
;;;           deepest inside the parcel (das-geom-pole) unless the centroid
;;;           is nearly as deep; real size taken from the bounding box and
;;;           reduced (down to 0.8x text, along the parcel then horizontal)
;;;           until the box plus a margin lies inside the parcel; finally
;;;           moved so the box centre sits on the anchor.
;;;
;;;   CHANGELOG 10.6.1 -> 10.7
;;;   [FIX-22] Default sizes came from fixed numbers (text 1.0, neighbour
;;;           text 2.5 at 3.0, length offset 2x text) or from whatever the
;;;           last parcel in the session used, so a 10 m parcel got 1 m
;;;           digits and a 3 m area label. A parcel without saved settings
;;;           (or with the untouched old defaults) now starts from its own
;;;           extent: text = bbox diagonal / 60 rounded to 1-1.25-1.5-2-
;;;           2.5-3-4-5-6-7.5 x 10^n; length gap 0.5x, marker 0.3x, small-
;;;           edge threshold 3x, neighbour text 1x at 2.5x. Precision,
;;;           marker style, mask and table columns still carry over.
;;;   [NEW-16] Boundary repair before QC (das-fix-run), after one question:
;;;           vertices closer than 1 cm to the previous one, a last vertex
;;;           on the first, an open polyline and a non-zero elevation are
;;;           fixed on the LWPOLYLINE itself (U undoes it with the rest).
;;;           Crossings and touches are not repaired; QC still reports them.
;;;
;;;   CHANGELOG 10.6 -> 10.6.1
;;;   [FIX-21] Lengths were drawn in the current DIMSTYLE, so a drawing
;;;           started from acad.dwt / acadiso.dwt got arrowheads, a
;;;           dimension line and extension lines over every parcel edge.
;;;           New DIMSTYLE DAS_LEN (das-dim-style-ensure): no arrows, no
;;;           dimension or extension lines, text DAS_ROMANS centred and
;;;           parallel to the edge, outside the parcel. The dialog's
;;;           offset (eb_offset, relabelled "Length Text Offset") is now the
;;;           clear gap between the edge and the near side of the number,
;;;           for full-size and small-edge text alike; it is not overridden.
;;;           Built from temporary DIM* variables that are
;;;           restored, so the user's current style is untouched. If the
;;;           style cannot be built, the lines are suppressed per object.
;;;
;;;   CHANGELOG 10.5.1 -> 10.6
;;;   [NEW-14] DASUPDATE redraws edited parcels without the dialog. Only
;;;           the area label was live (a FIELD); vertex numbers, dimensions,
;;;           neighbour captions and the coordinate table kept the old
;;;           geometry. The parcel record (LDATA DAS_PROJECT) now also
;;;           holds the dialog's drawing settings (*das-render-vars*), so
;;;           the redraw matches the first one; DAS reloads them too, except
;;;           the Length column, which follows DAS / DASL. Pick parcels, or
;;;           Enter for every parcel in the current space whose fingerprint
;;;           changed. Parcels with geometry errors are circled on L-QC and
;;;           skipped; neighbours are dropped when the vertex count changed.
;;;           The final-output path of DAS moved to das-output-final and
;;;           is shared with DASUPDATE (das-base-verts, das-start-verts).
;;;           Excel files are not regenerated.
;;;   [NEW-15] Warnings and questions no longer only scroll past on the
;;;           command line. das-ui-message is a small DCL window with any
;;;           buttons (alert has OK only): geometry errors (Continue/Stop),
;;;           geometry warnings, Excel file exists (Overwrite/New/Cancel)
;;;           and "export to Excel?" use it. Notes collected during a
;;;           command (das-note: neighbours dropped, Excel skipped, mask
;;;           failed, per-parcel results of DASUPDATE / DASQC / DASCHECK)
;;;           are shown together in one window at the end; each is still
;;;           printed, so the command history (F2) keeps them.
;;;           The window and the command line both lay text out left to
;;;           right, so the Farsi messages carry no handles and at most one
;;;           Latin word, placed with das-fa-end; Farsi lines are
;;;           right-aligned and end in an RLM (see das-rtl-p).
;;;   [FIX-20] Geometry QC missed a boundary that touches itself: a vertex
;;;           dropped onto another vertex or onto another side, or a side
;;;           running back along another one, pinches the parcel in two but
;;;           does not cross, so it got at most a "near vertices" warning.
;;;           New check 'touch (ERROR): a vertex closer than 1 cm to a side
;;;           it does not belong to (das-qc-touch-points, X-sorted sweep).
;;;           It replaces 'near, which it covers.
;;;
;;;   CHANGELOG 10.5 -> 10.5.1
;;;   [PERF-2] (nth i lst) walks i list cells, and the geometry code called
;;;           it inside per-vertex loops. The near-vertex check in das-qc-run
;;;           did so inside a double loop - O(n^3), run before the dialog
;;;           opens. It is now an X-sorted sweep (das-qc-near-pairs), and
;;;           the point-only checks moved to das-qc-check-pts. The
;;;           self-intersection test (all pairs of edges, the real bulk of
;;;           the QC time) now only compares edges whose bounding boxes
;;;           overlap, again via an X-sorted sweep, and still reports the
;;;           same first pair. Signed area, centroid, point-in-polygon,
;;;           label point, corner filter, start vertex, rotate, vertex
;;;           read-out, markers, dimensions, table, neighbour captions and
;;;           the longest edge now walk the list (das-list-next / -prev /
;;;           -take / -drop). QC on a 610-point boundary: 4.8 s -> 0.17 s.
;;;           Results are unchanged: checked against 10.5 in accoreconsole
;;;           on 9 generated parcels (297 comparisons, numbers and drawn
;;;           entities). das-draw-dimensions is ActiveX-only and was not
;;;           part of that run.
;;;
;;;   CHANGELOG 10.4.1 -> 10.5
;;;   [FIX-18] Applicant, national ID and building area were session
;;;           globals, so the next parcel's dialog opened with the
;;;           previous parcel's owner - easy to export a certificate in the
;;;           wrong name. Project data now lives on the polyline itself
;;;           (LDATA "DAS_PROJECT", MODULE 9B) together with the start
;;;           number, start rule and neighbour list; a polyline with no
;;;           record opens with empty applicant fields. The record holds
;;;           the owner's handle, so a COPY of the parcel starts empty.
;;;   [FIX-19] SaveAs ran with DisplayAlerts off and silently replaced an
;;;           existing Gvahi_Azla_/ParcelMap_ file of the same applicant.
;;;           Now asks Overwrite / New (_2, _3 ...) / Cancel.
;;;   [NEW-12] National ID is checked: 10-digit personal code (check
;;;           digit mod 11, 8-9 digits are zero-padded) or 11-digit legal
;;;           entity ID. Persian/Arabic digits, blanks and dashes are
;;;           normalised. A failing ID blocks OK once; pressing OK again
;;;           keeps it (foreign nationals, unusual codes). The Excel
;;;           export is only offered when applicant and ID are filled in.
;;;   [NEW-13] Every entity DAS draws is tagged with XDATA "DAS_OUT"
;;;           (1005 = parcel handle). Running DAS again on the same parcel
;;;           deletes its previous output instead of drawing a second
;;;           copy on top. Output from 10.4.1 and earlier is untagged and
;;;           has to be erased by hand once.
;;;
;;;   CHANGELOG 10.4 -> 10.4.1
;;;   [NEW-11] The Excel templates temp1.xlsx / temp2.xlsx now live in the
;;;           repository (templates\) and are found by das-find-template:
;;;           saved folder, support paths, templates\ under a support
;;;           path, then C:\DAS_Templates. If missing, the user is asked
;;;           once and the folder is remembered (setting DAS_TEMPLATE_DIR).
;;;
;;;   CHANGELOG 10.3 -> 10.4
;;;   [NEW-10] DASIMP (MODULE 10): imports points from a CSV/TXT file.
;;;           Delimiter (tab ; , or blanks) and header are detected
;;;           automatically; a UTF-8 BOM is stripped; decimal commas are
;;;           accepted. Columns are mapped from the header names
;;;           (Name/X/Y/Z/Code, E/N, East/North...) or, without a header,
;;;           from the data itself, and X/Y are swapped when the "X"
;;;           column holds northings. Each point is classified by the
;;;           prefix of its name (B CP BL W R S C P X), then by the Code
;;;           column, and the rest are assigned in one question. Before
;;;           anything is drawn the command prints a summary (counts per
;;;           prefix, duplicate names/coordinates, points outside the UTM
;;;           range of Iran) and shows the points on screen with grdraw;
;;;           Yes / No / Map (pick columns by hand) / IncludeX. Drawn as
;;;           POINT on L-IMP-<prefix> with the name as TEXT on
;;;           L-IMP-<prefix>-T. X (rejected) points are skipped unless
;;;           IncludeX is chosen.
;;;
;;;   CHANGELOG 10.2.2 -> 10.3
;;;   [NEW-7] Geometry QC (MODULE 3C). DAS/DASL check the polyline before
;;;           the dialog: arc segments, self-crossing, open polyline,
;;;           near-duplicate vertices, sides < 5 cm, spikes < 5 deg,
;;;           non-zero Z, non-+Z normal (OCS) and coordinates outside the
;;;           UTM range for Iran. Problems are circled in red on layer
;;;           L-QC. Errors stop DAS unless the user confirms. New command
;;;           DASQC runs the same checks on many polylines at once.
;;;   [NEW-8] The area label is a live FIELD bound to the polyline's
;;;           Area, so it updates when the parcel is edited. Falls back
;;;           to static text if the field cannot be created.
;;;   [NEW-9] Change tracking (MODULE 9). The final output stamps the
;;;           polyline with XDATA (app DAS_TOOLS): version, date, vertex
;;;           count, area and a coordinate checksum. New command
;;;           DASCHECK lists parcels edited after their output was made.
;;;   [FIX-17] das-geom-get-vertices read LWPOLYLINE coordinates in OCS.
;;;            A polyline that was mirrored or drawn in a tilted UCS has a
;;;            normal other than +Z, and its X came out wrong (mirrored).
;;;            Coordinates are now converted to WCS with trans.
;;;
;;;   CHANGELOG 10.2.1 -> 10.2.2
;;;   [DEL] Edge-walk neighbour entry ([NEW-6]) removed: the Walk/Dialog
;;;         prompt before the dialog, das-ui-walk-neighbors and
;;;         das-walk-highlight-run. Neighbours are entered in the dialog
;;;         only. das-nb-compose is kept (used by the dialog "+ Add").
;;;
;;;   CHANGELOG 10.2 -> 10.2.1
;;;   Added das-pt-in-poly-p, das-safe-label-point,
;;;   das-angle-is-steep-p, das-walk-highlight-run.
;;;
;;;   CHANGELOG 10.1 -> 10.2
;;;   [NEW-6] Edge-walk neighbour entry. c:das / c:dasl now offer to walk
;;;           the parcel's real SIDES on screen and ask for each one's
;;;           neighbour, instead of typing point-number ranges into the
;;;           dialog by hand. Sides are grouped by actual corner (reusing
;;;           the [NEW-4] corner filter), not by raw edge - a
;;;           GPS-collected boundary can carry many nearly-straight
;;;           points along one physical side, and asking a separate
;;;           question for each of those made the walk needlessly slow
;;;           on high-vertex parcels. Enter repeats the last neighbour
;;;           typed and extends its range - so several consecutive sides
;;;           along one alley become a single table row, not one row
;;;           each. "0" marks a side as having no neighbour. Each side
;;;           highlights with grdraw as you walk it; nothing is
;;;           entmake'd, so there is nothing to clean up if you cancel
;;;           (Esc) partway through - whatever ranges were already
;;;           confirmed stay, and the dialog's own "+ Add" list is still
;;;           there afterwards to fix or finish the rest by hand.
;;;           Text composition (both the stored data and the list-box
;;;           string) was factored out of das-ui-add-neighbor into
;;;           das-nb-compose, so the dialog path and the walk path can
;;;           never disagree on wording for the same input. See notes at
;;;           das-ui-walk-neighbors.
;;;   [FIX-14] das-draw-range-neighbors placed a caption relative to the
;;;            CHORD between a range's first and last point. On a run
;;;            that bends or arcs, the real boundary could bulge past
;;;            that chord on the caption's side, and a fixed offset was
;;;            not always enough to clear it - the text ended up sitting
;;;            on top of the polyline. Now measures how far the farthest
;;;            in-range vertex already bulges past the chord and adds
;;;            that on top of the normal offset.
;;;   [FIX-15] The area label used the bounding-box centre, which on a
;;;            concave/notched parcel routinely falls outside the
;;;            boundary entirely. Now tries the true area centroid
;;;            (das-geom-centroid, already used for the lat/lon report)
;;;            first, and only if THAT also lands outside falls back to
;;;            das-safe-label-point - the midpoint of the widest
;;;            horizontal span through the centroid's own height,
;;;            guaranteed to be inside.
;;;   [FIX-16] The area label's rotation followed whichever edge was
;;;            longest, which could leave it steep/near-vertical - hard
;;;            to read for a summary label that (unlike a dimension) has
;;;            no single edge it must align with. Snaps to horizontal
;;;            whenever the longest edge is closer to vertical than
;;;            horizontal (das-angle-is-steep-p).
;;;
;;;   NOT YET DONE (next step): layout/title-block automation (DASLAY).
;;; ============================================================

;;; ============================================================
;;; MODULE 1: GLOBALS
;;; ============================================================

(setq *das-version* "10.7.3")

;; [FIX-25] Farsi text can fail to show ("?") for two reasons, and DAS
;; then runs with the English interface instead:
;;  1. AutoCAD before 2021 reads this file byte by byte (LISP text is not
;;     Unicode there), so every Persian letter becomes two characters.
;;     das-unicode-p reads a Persian letter back: one character code
;;     means Unicode strings work.
;;  2. Dialogs, alerts and the command line also go through the Windows
;;     ANSI code page ("Language for non-Unicode programs"). Only 1256
;;     (Persian/Arabic) holds Persian letters. The code page is read from
;;     the registry; if it cannot be read, Farsi is allowed.
(defun das-unicode-p ()
  (equal (vl-string->list "ب") '(1576)))

(defun das-acp ( / v)
  (setq v (vl-catch-all-apply 'vl-registry-read
            (list "HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\Nls\\CodePage" "ACP")))
  (if (= (type v) 'STR) v nil))

(defun das-farsi-ok-p ()
  (and (das-unicode-p) (member (das-acp) '(nil "1256")) T))

(defun das-farsi-why ()
  (if (not (das-unicode-p))
    "this AutoCAD is older than 2021"
    (strcat "the Windows language for non-Unicode programs is code page "
            (das-acp) ", not Persian (1256)")))

;; chr for codes above 255 (direction marks); empty where strings are not Unicode.
(defun das-uchr (n) (if (das-unicode-p) (chr n) ""))

(defun das-init-globals ()
  (if (null *das-last-scale*)     (setq *das-last-scale* 1.0))
  (if (null *das-last-start*)     (setq *das-last-start* 1))
  (if (null *das-last-offset*)    (setq *das-last-offset* nil))
  (if (null *das-last-prec-c*)    (setq *das-last-prec-c* 4))
  (if (null *das-last-prec-l*)    (setq *das-last-prec-l* 2))
  (if (null *das-last-prec-a*)    (setq *das-last-prec-a* 2))

  (if (null *das-last-mask*)      (setq *das-last-mask* "No"))
  (if (null *das-last-mkstyle*)   (setq *das-last-mkstyle* "Circle"))
  (if (null *das-last-mksize*)    (setq *das-last-mksize* nil))
  (if (null *das-last-small-th*)  (setq *das-last-small-th* 0.0))
  (if (null *das-last-small-fac*) (setq *das-last-small-fac* 0.7))
  (if (null *das-last-col-count*) (setq *das-last-col-count* 1))

  (if (null *das-last-nb-scale*)  (setq *das-last-nb-scale* 2.5))
  (if (null *das-last-nb-offset*) (setq *das-last-nb-offset* 3.0))

  (if (null *das-last-withlen*)   (setq *das-last-withlen* nil))
  (if (null *das-lang*)           (setq *das-lang* "Farsi"))
  (if (not (das-farsi-ok-p))      (setq *das-lang* "English"))   ; [FIX-25]

  ;; "NorthWest" | "Northernmost" | "AsDrawn"
  (if (null *das-start-rule*)     (setq *das-start-rule* "NorthWest"))

  ;; [NEW-4] Vertices whose interior angle exceeds this are treated as
  ;; lying on a straight side and cannot become vertex number 1.
  ;; This MUST equal max_angle in cadastre.py on utmdoc.ir, and
  ;; START_VERTEX_RULE there MUST be nw_metric, or the web service and
  ;; DAS will number the same parcel differently.
  (if (null *das-corner-max-angle*) (setq *das-corner-max-angle* 175.0))

  ;; Excel / cadastral project data
  ;; [FIX-18] these are per parcel - das-proj-load fills them from the
  ;; polyline, or empties them for a parcel that has no record yet
  (if (null *das-applicant*)      (setq *das-applicant* ""))
  (if (null *das-nat-id*)         (setq *das-nat-id* ""))
  (if (null *das-b-area*)         (setq *das-b-area* 0.0))
  (if (null *das-zone*)           (setq *das-zone* 39))

  ;; [NEW-10] DASIMP: last file, label height, and the prefix table
  ;; (prefix layer colour title). The prefix is the leading letters of the
  ;; point name ("B12" -> B, "CP3" -> CP), else of the Code column.
  (if (null *das-imp-texth*)      (setq *das-imp-texth* 1.0))
  (if (null *das-imp-last-path*)  (setq *das-imp-last-path* ""))
  (setq *das-imp-cats*
    '(("B"  "L-IMP-B"  1  "مرز")
      ("CP" "L-IMP-CP" 5  "نقطهٔ کنترل")
      ("BL" "L-IMP-BL" 30 "بنا")
      ("W"  "L-IMP-W"  8  "دیوار")
      ("R"  "L-IMP-R"  2  "معبر")
      ("S"  "L-IMP-S"  6  "سوله")
      ("C"  "L-IMP-C"  4  "کانال")
      ("P"  "L-IMP-P"  3  "عارضهٔ نقطه‌ای")
      ("X"  "L-IMP-X"  9  "مردود")))

  (setq *das-layers*
    (list (list "L-POINTS"   1)
          (list "L-DIM"      8)
          (list "L-TEXT"     5)
          (list "L-TABLE"    7)
          (list "L-AREA"     7)   ;; [FIX-23] was 2 (yellow)
          (list "L-LEADER"   6)
          (list "L-NEIGHBOR" 7)))   ;; [FIX-10] was 3 (green); black reads better in print
  (princ)
)
(das-init-globals)

;;; ============================================================
;;; MODULE 2: UTILITIES
;;; ============================================================

;; [FIX-3] numeric guards ------------------------------------------------
(defun das-clamp (v lo hi)
  (cond ((null v) lo)
        ((< v lo) lo)
        ((> v hi) hi)
        (T v)))

(defun das-clamp-int (v lo hi) (fix (das-clamp (float v) (float lo) (float hi))))

;; [FIX-7] strip characters Windows will not accept in a filename -------
(defun das-safe-filename (s / bad out i c)
  (if (or (null s) (= s "")) (setq s "unnamed"))
  (setq bad "\\/:*?\"<>|\t\n" out "" i 0)
  (repeat (strlen s)
    (setq i (1+ i) c (substr s i 1))
    (setq out (strcat out (if (vl-string-search c bad) "_" c))))
  (setq out (vl-string-trim " ." out))
  (if (= out "") "unnamed" out))

;; [FIX-8] entmake-based layer creation ---------------------------------
;; The old version drove the LAYER command, which echoes to the command
;; line, cannot run transparently, and fails if called while another
;; command is active.
(defun das-utils-ensure-layer (layName color / ent dxf)
  (if (setq ent (tblobjname "layer" layName))
    (progn
      (setq dxf (entget ent))
      ;; clear the Frozen (1) and Locked (4) bits, force colour ON
      (setq dxf (subst (cons 70 (logand (cdr (assoc 70 dxf)) (~ 5))) (assoc 70 dxf) dxf))
      (setq dxf (subst (cons 62 (abs color)) (assoc 62 dxf) dxf))
      (vl-catch-all-apply 'entmod (list dxf)))
    (entmake (list '(0 . "LAYER")
                   '(100 . "AcDbSymbolTableRecord")
                   '(100 . "AcDbLayerTableRecord")
                   (cons 2 layName) (cons 70 0)
                   (cons 62 color) (cons 6 "Continuous"))))
  layName)

;; [FIX-8] entmake-based style creation ---------------------------------
;; -STYLE asks "Vertical? <N>" for SHX fonts but NOT for TTF fonts, so
;; the old fixed answer list ("N" "N" "N") sent one surplus keystroke
;; into whatever command ran next when creating DAS_FARSI.
(defun das-utils-ensure-style (styleName fontFile / ent dxf)
  (if (setq ent (tblobjname "style" styleName))
    (progn
      (setq dxf (entget ent))
      ;; a fixed text height (DXF 40) makes das-utils-text-center's
      ;; height argument silently ignored, so force it back to 0
      (if (/= (cdr (assoc 40 dxf)) 0.0)
        (progn
          (setq dxf (subst (cons 40 0.0) (assoc 40 dxf) dxf))
          (vl-catch-all-apply 'entmod (list dxf)))))
    (entmake (list '(0 . "STYLE")
                   '(100 . "AcDbSymbolTableRecord")
                   '(100 . "AcDbTextStyleTableRecord")
                   (cons 2 styleName) (cons 70 0)
                   (cons 40 0.0) (cons 41 1.0) (cons 50 0.0)
                   (cons 71 0) (cons 42 2.5)
                   (cons 3 fontFile) (cons 4 ""))))
  styleName)

(defun das-utils-mkline (p1 p2 layer)
  (entmakex (list (cons 0 "LINE") (cons 8 layer) (cons 62 256)
                  (cons 10 p1) (cons 11 p2))))

(defun das-utils-text-center (pt hgt str layer styleName)
  (entmakex
    (list (cons 0 "TEXT") (cons 8 layer) (cons 7 styleName)
          (cons 10 pt) (cons 11 pt) (cons 40 hgt) (cons 1 str)
          (cons 50 0.0) (cons 72 1) (cons 73 2) (cons 62 256))))

;; [FIX-11] A TEXT entity cannot render a superscript, so the area label
;; used to read "m2". MTEXT can, via the stacking code \S<up>^<down>; -
;; writing "m\S2^;" puts the 2 above the baseline with nothing under it.
;; Attachment 5 is middle-centre, matching the old TEXT alignment (72/73).
(defun das-utils-mtext-rotated (pt hgt str layer styleName ang)
  (entmakex
    (list (cons 0 "MTEXT") (cons 100 "AcDbEntity") (cons 8 layer)
          (cons 100 "AcDbMText") (cons 10 pt) (cons 40 hgt)
          (cons 1 str) (cons 7 styleName) (cons 50 ang)
          (cons 71 5) (cons 62 256))))

;; [NEW-8] Area label as a live FIELD bound to the polyline's Area, so
;; editing the parcel after DAS no longer leaves a stale number. Field
;; codes are only evaluated when the MText is created through ActiveX;
;; entmake would store them as plain text. Returns the ename, or nil if
;; anything fails, in which case the caller draws static text instead.
(defun das-utils-objid-string (obj / util)
  (setq util (vla-get-Utility (vla-get-ActiveDocument (vlax-get-acad-object))))
  (if (vlax-method-applicable-p util 'GetObjectIdString)
    (vla-GetObjectIdString util obj :vlax-false)
    (itoa (vla-get-ObjectID obj))))

(defun das-utils-area-field-mtext (pt hgt obj prec layer styleName ang / res)
  (setq res
    (vl-catch-all-apply
      (function
        (lambda ( / doc spc str mt)
          (setq doc (vla-get-ActiveDocument (vlax-get-acad-object))
                spc (vl-catch-all-apply 'vla-ObjectIDToObject
                                        (list doc (vla-get-OwnerID obj))))
          (if (vl-catch-all-error-p spc) (setq spc (vla-get-ModelSpace doc)))
          (setq str (strcat "S=%<\\AcObjProp Object(%<\\_ObjId "
                            (das-utils-objid-string obj)
                            ">%).Area \\f \"%lu2%pr" (itoa prec) "\">% m\\S2^;"))
          (setq mt (vla-AddMText spc (vlax-3d-point pt) 0.0 str))
          (vla-put-AttachmentPoint mt acAttachmentPointMiddleCenter)
          (vla-put-InsertionPoint mt (vlax-3d-point pt))
          (vla-put-Height mt hgt)
          (vla-put-Rotation mt ang)
          (vla-put-Layer mt layer)
          (vla-put-StyleName mt styleName)
          (vlax-vla-object->ename mt)))))
  (if (vl-catch-all-error-p res) nil res))

(defun das-utils-text-rotated (pt hgt str layer styleName ang)
  (entmakex
    (list (cons 0 "TEXT") (cons 8 layer) (cons 7 styleName)
          (cons 10 pt) (cons 11 pt) (cons 40 hgt) (cons 1 str)
          (cons 50 ang) (cons 72 1) (cons 73 2) (cons 62 256))))

;; [FIX-3] entmakex returns nil when the text height is 0 or the string
;; is empty. (entget nil) then raised "bad argument type".
;; [FIX-13] The mask used to shell out to "_.-TEXTMASK". That command
;; does not exist - Express Tools defines TEXTMASK with no leading
;; hyphen, its prompt sequence changes between releases, and it is not
;; installed on every machine (LT has no Express Tools at all). The old
;; guard (boundp 'c:textmask) passed while the command still failed.
;;
;; Native replacement: put a WIPEOUT rectangle behind each label, then
;; bring the labels to the front. Draw order ends up lines -> wipeouts ->
;; text, which is what a mask needs. No external dependency.
;;
;; The rectangle comes from vla-GetBoundingBox, which is axis-aligned.
;; That is exact for the table because its text is horizontal. Rotated
;; text would get an oversized box, so only pass horizontal labels here.
(defun das-utils-apply-mask (textList offset / e obj mn mx r
                             p1 p2 p3 p4 made ss oldCmd frameVal)
  (setq made '())
  (if textList
    (progn
      (setq oldCmd (getvar 'cmdecho))
      (setvar 'cmdecho 0)

      (foreach e textList
        (if (and e (entget e))
          (progn
            (setq obj (vl-catch-all-apply 'vlax-ename->vla-object (list e)))
            (if (not (vl-catch-all-error-p obj))
              (progn
                (setq r (vl-catch-all-apply 'vla-GetBoundingBox (list obj 'mn 'mx)))
                (if (not (vl-catch-all-error-p r))
                  (progn
                    (setq mn (vlax-safearray->list mn)
                          mx (vlax-safearray->list mx))
                    (setq p1 (list (- (car mn) offset) (- (cadr mn) offset) 0.0)
                          p3 (list (+ (car mx) offset) (+ (cadr mx) offset) 0.0)
                          p2 (list (car p3) (cadr p1) 0.0)
                          p4 (list (car p1) (cadr p3) 0.0))
                    (setq r (vl-catch-all-apply 'vl-cmdf
                              (list "_.WIPEOUT" "_non" p1 "_non" p2
                                                "_non" p3 "_non" p4 "")))
                    (if (not (vl-catch-all-error-p r))
                      (setq made (cons (entlast) made))))))))))

      ;; lift the labels above the masks just created
      (if made
        (progn
          (setq ss (ssadd))
          (foreach e textList (if (and e (entget e)) (ssadd e ss)))
          (if (> (sslength ss) 0)
            (vl-catch-all-apply 'vl-cmdf (list "_.DRAWORDER" ss "" "_F")))))

      ;; WIPEOUTFRAME is read when the drawing is displayed, not when the
      ;; wipeout is made, so saving and restoring it around the loop would
      ;; achieve nothing. Set it once and say so.
      (if made
        (progn
          (setq frameVal (vl-catch-all-apply 'getvar (list "WIPEOUTFRAME")))
          (if (and (not (vl-catch-all-error-p frameVal)) (/= frameVal 0))
            (progn
              (vl-catch-all-apply 'setvar (list "WIPEOUTFRAME" 0))
              (princ (das-tr (strcat "\n" (das-fa-end "برای پنهان شدن کادر ماسک‌ها، این متغیر روی صفر گذاشته شد:" "WIPEOUTFRAME"))
                             "\nWIPEOUTFRAME was set to 0 so the mask frames stay hidden."))))))

      (setvar 'cmdecho oldCmd)
      (if (null made)
        (das-note (das-tr (strcat "\n" (das-fa-end "هشدار: ماسک جدول ساخته نشد، چون این فرمان در دسترس نبود:" "WIPEOUT"))
                          "\nWarning: the table mask was not created (WIPEOUT was not available)."))))
  )
  (reverse made)
)

;; [PERF-2] list helpers. (nth i lst) walks i conses, so a loop that
;; calls it for every vertex is quadratic and one that does so inside a
;; second loop is cubic. The geometry code walks these instead:
;; das-list-next is lst shifted so element i is vertex i+1 (wrapping),
;; das-list-prev is the same towards i-1.
(defun das-list-next (lst) (if lst (append (cdr lst) (list (car lst)))))
(defun das-list-prev (lst) (if lst (cons (last lst) (reverse (cdr (reverse lst))))))

(defun das-list-take (lst n / out)
  (setq out '())
  (while (and lst (> n 0))
    (setq out (cons (car lst) out) lst (cdr lst) n (1- n)))
  (reverse out))

(defun das-list-drop (lst n)
  (while (and lst (> n 0)) (setq lst (cdr lst) n (1- n)))
  lst)

(defun das-utils-remove-at (idx lst / i new_lst)
  (setq i 0 new_lst '())
  (foreach x lst
    (if (/= i idx) (setq new_lst (cons x new_lst)))
    (setq i (1+ i)))
  (reverse new_lst))

;; [FIX-12] Collect the entities created after a marker entity, and take
;; their combined bounding box. The coordinate table used to be placed
;; 5 units below the polygon itself, which ignored the neighbour text
;; drawn outside the boundary - on parcels with long neighbour captions
;; the two overlapped. Now the table is placed below everything actually
;; drawn. Entities that refuse GetBoundingBox are skipped rather than
;; aborting the whole run.
(defun das-ents-after (marker / e out)
  (setq out '() e (if marker (entnext marker) (entnext)))
  (while e (setq out (cons e out) e (entnext e)))
  (reverse out))

(defun das-bbox-of (ents / e obj mn mx bx minx miny maxx maxy)
  (setq minx nil)
  (foreach e ents
    (setq obj (vl-catch-all-apply 'vlax-ename->vla-object (list e)))
    (if (not (vl-catch-all-error-p obj))
      (progn
        (setq bx (vl-catch-all-apply 'vla-GetBoundingBox (list obj 'mn 'mx)))
        (if (not (vl-catch-all-error-p bx))
          (progn
            (setq mn (vlax-safearray->list mn) mx (vlax-safearray->list mx))
            (if (null minx)
              (setq minx (car mn) miny (cadr mn) maxx (car mx) maxy (cadr mx))
              (setq minx (min minx (car mn)) miny (min miny (cadr mn))
                    maxx (max maxx (car mx)) maxy (max maxy (cadr mx)))))))))
  (if minx (list minx miny maxx maxy) nil))

(defun das-clean-preview (ents / e)
  (foreach e ents (if (and e (entget e)) (entdel e)))
  (princ))

;;; ============================================================
;;; MODULE 3: GEOMETRY ENGINE
;;; ============================================================

(defun das-geom-get-vertices (ent / obj coords ed)
  (setq obj    (vlax-ename->vla-object ent)
        coords (vlax-get obj 'Coordinates)
        ed     (entget ent))
  (das-geom-coords->pts coords ed ent))

;; [PERF-2] split out of das-geom-get-vertices so it can be tested
;; without ActiveX; walks the flat list instead of (nth i coords).
(defun das-geom-coords->pts (coords ed ent / vertices elev ocs p)
  (setq vertices '())
  (cond
    ((= (cdr (assoc 0 ed)) "LWPOLYLINE")
     ;; [FIX-17] LWPOLYLINE coordinates are in the entity's OCS. With a
     ;; normal other than +Z (mirrored, or drawn in a tilted UCS) they
     ;; are not WCS coordinates, so convert them.
     (setq elev (cond ((cdr (assoc 38 ed))) (0.0))
           ocs  (not (equal (cond ((cdr (assoc 210 ed))) ('(0.0 0.0 1.0)))
                            '(0.0 0.0 1.0) 1e-9)))
     (while (cdr coords)
       (setq p (list (car coords) (cadr coords) elev) coords (cddr coords))
       (if ocs (setq p (trans p ent 0)))
       (setq vertices (cons (list (car p) (cadr p) 0.0) vertices))))
    (T
     (while (cddr coords)
       (setq vertices (cons (list (car coords) (cadr coords) (caddr coords)) vertices)
             coords   (cdddr coords)))))
  (reverse vertices))

;; Reject repeated vertices. A duplicated point produces a zero-length
;; edge in the table and in both Excel sheets.
(defun das-geom-dedupe (pts tol / out prev p)
  (setq out '() prev nil)
  (foreach p pts
    (if (or (null prev) (> (distance prev p) tol))
      (progn (setq out (cons p out)) (setq prev p))))
  (setq out (reverse out))
  ;; also drop a closing vertex identical to the first
  (if (and (> (length out) 2)
           (<= (distance (car out) (last out)) tol))
    (setq out (reverse (cdr (reverse out)))))
  out)

(defun das-geom-signed-area (pts / a b s)
  (setq a pts b (das-list-next pts) s 0.0)
  (while a
    (setq s (+ s (- (* (car (car a)) (cadr (car b)))
                    (* (car (car b)) (cadr (car a)))))
          a (cdr a) b (cdr b)))
  s)

(defun das-geom-calc-marker-size (vertices / minx maxx miny maxy diagLen p)
  (setq minx (car (car vertices))  maxx (car (car vertices))
        miny (cadr (car vertices)) maxy (cadr (car vertices)))
  (foreach p vertices
    (setq minx (min minx (car p)) maxx (max maxx (car p))
          miny (min miny (cadr p)) maxy (max maxy (cadr p))))
  (setq diagLen (distance (list minx miny) (list maxx maxy)))
  (max 0.05 (min 1.50 (* diagLen 0.005))))

;; [FIX-22] Every size in the dialog used to start from fixed numbers
;; (text 1.0, neighbours 2.5 / 3.0, length offset 2.0) or from the last
;; parcel drawn in the session. On a 10 m parcel that meant 1 m digits,
;; a 3 m area label and neighbour captions far outside the sheet; on a
;; large farm the same numbers were too small. A parcel without its own
;; saved settings now starts from sizes derived from its own extent:
;; text height = bounding-box diagonal / 60, rounded to a "nice" value,
;; and every other size a fixed multiple of it. Settings saved on the
;; parcel are kept, except the untouched pre-10.7 defaults (das-sizes-legacy-p).
(defun das-nice-num (x / e m)
  (if (<= x 0.0)
    0.1
    (progn
      (setq e (expt 10.0 (fix (if (< x 1.0) (1- (/ (log x) (log 10.0))) (/ (log x) (log 10.0)))))
            m (/ x e))
      (* e (cond ((< m 1.12) 1.0) ((< m 1.37) 1.25) ((< m 1.75) 1.5) ((< m 2.25) 2.0)
                 ((< m 2.75) 2.5) ((< m 3.5) 3.0) ((< m 4.5) 4.0) ((< m 5.5) 5.0)
                 ((< m 6.75) 6.0) ((< m 8.75) 7.5) (T 10.0))))))

(defun das-auto-text-height (vertices / minx maxx miny maxy p)
  (setq minx (car (car vertices))  maxx minx
        miny (cadr (car vertices)) maxy miny)
  (foreach p vertices
    (setq minx (min minx (car p)) maxx (max maxx (car p))
          miny (min miny (cadr p)) maxy (max maxy (cadr p))))
  (max 0.01 (das-nice-num (/ (distance (list minx miny) (list maxx maxy)) 60.0))))

(defun das-auto-sizes (vertices / h)
  (setq h (das-auto-text-height vertices))
  (setq *das-last-scale*     h                 ; vertex numbers, lengths, table
        *das-last-offset*    (* h 0.5)         ; gap edge -> length text
        *das-last-mksize*    (* h 0.3)         ; vertex marker
        *das-last-small-th*  (* h 3.0)         ; shorter edges get smaller text
        *das-last-nb-scale*  h                 ; neighbour captions
        *das-last-nb-offset* (* h 2.5))        ; outside the length text
  h)

;; The record's settings are the dialog's untouched 10.6-and-earlier
;; defaults, i.e. nobody chose them for this parcel.
(defun das-sizes-legacy-p (settings / g)
  (defun g (k) (cadr (assoc k settings)))
  (and (numberp (g "*DAS-LAST-SCALE*"))     (equal (g "*DAS-LAST-SCALE*") 1.0 1e-9)
       (numberp (g "*DAS-LAST-NB-SCALE*"))  (equal (g "*DAS-LAST-NB-SCALE*") 2.5 1e-9)
       (numberp (g "*DAS-LAST-NB-OFFSET*")) (equal (g "*DAS-LAST-NB-OFFSET*") 3.0 1e-9)))

(defun das-sizes-need-auto-p (ent / d st)
  (setq d (das-proj-read ent) st (if d (das-proj-get d "settings")))
  (or (null st) (das-sizes-legacy-p st)))

(defun das-geom-centroid (pts / a b cx cy signed-area pt1 pt2 factor)
  (setq a pts b (das-list-next pts) cx 0.0 cy 0.0)
  (setq signed-area (/ (das-geom-signed-area pts) 2.0))
  (while a
    (setq pt1 (car a) pt2 (car b) a (cdr a) b (cdr b))
    (setq factor (- (* (car pt1) (cadr pt2)) (* (car pt2) (cadr pt1))))
    (setq cx (+ cx (* (+ (car pt1) (car pt2)) factor)))
    (setq cy (+ cy (* (+ (cadr pt1) (cadr pt2)) factor))))
  (if (zerop signed-area)
    (list (car (nth 0 pts)) (cadr (nth 0 pts)))
    (list (/ cx (* 6.0 signed-area)) (/ cy (* 6.0 signed-area)))))

;; [FIX-15] das-geom-centroid is the true area centroid, but on a concave
;; parcel (L-shaped, U-shaped, anything with a notch) the centroid can
;; still land outside the boundary - the area label used to be placed at
;; the bounding-box centre instead, which is even more likely to miss.
;; das-pt-in-poly-p is a standard even-odd ray cast; das-safe-label-point
;; is the fallback when the centroid fails it: a horizontal line through
;; the centroid's Y crosses the boundary an even number of times, and the
;; midpoint of the WIDEST of those crossing-to-crossing spans is
;; guaranteed to sit inside the polygon (as long as that line crosses it
;; at all, true for every simple polygon at its own centroid height).
(defun das-pt-in-poly-p (pt pts / pa pb xi yi xj yj px py inside)
  (setq pa pts pb (das-list-next pts) inside nil px (car pt) py (cadr pt))
  (while pa
    (setq xi (car (car pa)) yi (cadr (car pa))
          xj (car (car pb)) yj (cadr (car pb))
          pa (cdr pa) pb (cdr pb))
    (if (and (/= (>= yi py) (>= yj py))
             (< px (+ xi (* (/ (- py yi) (- yj yi)) (- xj xi)))))
      (setq inside (not inside))))
  inside)

(defun das-safe-label-point (centroid pts / pa pb y0 xi yi xj yj xs best bestw a b w)
  (setq pa pts pb (das-list-next pts) y0 (cadr centroid) xs '())
  (while pa
    (setq xi (car (car pa)) yi (cadr (car pa))
          xj (car (car pb)) yj (cadr (car pb))
          pa (cdr pa) pb (cdr pb))
    (if (/= (>= yi y0) (>= yj y0))
      (setq xs (cons (+ xi (* (/ (- y0 yi) (- yj yi)) (- xj xi))) xs))))
  (setq xs (vl-sort xs '<))
  (setq best (car centroid) bestw -1.0)
  (while (>= (length xs) 2)
    (setq a (car xs) b (cadr xs) w (- b a))
    (if (> w bestw) (setq bestw w best (/ (+ a b) 2.0)))
    (setq xs (cddr xs)))
  (list best y0))

;; [FIX-23] Area label placement. The label used to go to the centroid
;; at 3x the text height whatever the parcel's shape, so on a small or
;; narrow parcel it ran over the boundary. Now: the anchor is the point
;; farthest from every side (a grid search, refined twice - the "pole of
;; inaccessibility"), unless the centroid is nearly as deep inside; the
;; label's real size is measured from its bounding box; and the height is
;; reduced step by step (along the parcel, then horizontal) until the
;; whole box, plus a margin, lies inside the parcel.
(defun das-geom-seg-dist (p a b / dx dy l2 k q)
  (setq dx (- (car b) (car a)) dy (- (cadr b) (cadr a))
        l2 (+ (* dx dx) (* dy dy)))
  (if (< l2 1e-18)
    (distance (list (car p) (cadr p)) (list (car a) (cadr a)))
    (progn
      (setq k (/ (+ (* (- (car p) (car a)) dx) (* (- (cadr p) (cadr a)) dy)) l2)
            k (max 0.0 (min 1.0 k))
            q (list (+ (car a) (* k dx)) (+ (cadr a) (* k dy))))
      (distance (list (car p) (cadr p)) q))))

(defun das-geom-edge-dist (p pts / pa pb d)
  (setq pa pts pb (das-list-next pts) d 1e99)
  (while pa
    (setq d (min d (das-geom-seg-dist p (car pa) (car pb))) pa (cdr pa) pb (cdr pb)))
  d)

;; (point depth): the inside point farthest from the boundary.
(defun das-geom-pole (pts / minx maxx miny maxy best bd step i j q d cx cy span)
  (setq minx (car (car pts)) maxx minx miny (cadr (car pts)) maxy miny)
  (foreach q pts
    (setq minx (min minx (car q)) maxx (max maxx (car q))
          miny (min miny (cadr q)) maxy (max maxy (cadr q))))
  (setq best nil bd -1.0
        cx (/ (+ minx maxx) 2.0) cy (/ (+ miny maxy) 2.0)
        span (max (- maxx minx) (- maxy miny)))
  (repeat 3
    (setq step (/ span 14.0) i -7)
    (while (<= i 7)
      (setq j -7)
      (while (<= j 7)
        (setq q (list (+ cx (* i step)) (+ cy (* j step))))
        (if (and (das-pt-in-poly-p q pts) (> (setq d (das-geom-edge-dist q pts)) bd))
          (setq best q bd d))
        (setq j (1+ j)))
      (setq i (1+ i)))
    (if best (setq cx (car best) cy (cadr best)))
    (setq span (/ span 5.0)))
  (if best (list best bd) nil))

;; Does a w x h box centred on c, turned by ang, lie inside pts with
;; at least margin to every side? Corners and side midpoints must be
;; inside and clear of the sides, and no parcel vertex may poke into it.
(defun das-label-fits-p (c w h ang pts margin / ux uy vx vy hw hh probes ok q lx ly f)
  (setq ux (cos ang) uy (sin ang) vx (- uy) vy ux hw (/ w 2.0) hh (/ h 2.0) ok T)
  (foreach f '((-1 -1) (1 -1) (1 1) (-1 1) (0 -1) (0 1) (-1 0) (1 0))
    (setq probes (cons (list (+ (car c) (* (car f) hw ux) (* (cadr f) hh vx))
                             (+ (cadr c) (* (car f) hw uy) (* (cadr f) hh vy)))
                       probes)))
  (foreach q probes
    (if (or (not (das-pt-in-poly-p q pts)) (< (das-geom-edge-dist q pts) margin))
      (setq ok nil)))
  (if ok
    (foreach q pts
      (setq lx (+ (* (- (car q) (car c)) ux) (* (- (cadr q) (cadr c)) uy))
            ly (+ (* (- (car q) (car c)) vx) (* (- (cadr q) (cadr c)) vy)))
      (if (and (< (abs lx) (+ hw margin)) (< (abs ly) (+ hh margin))) (setq ok nil))))
  ok)

;; Size and place the area label ent (an MTEXT, middle-centre). h0 is the
;; preferred height, hMin the smallest acceptable one.
(defun das-area-label-place (ent pts centroid ang h0 hMin / obj pole c mn mx w0 t0 hh ok angs a margin ext)
  (setq obj (vlax-ename->vla-object ent))
  ;; anchor
  (setq pole (das-geom-pole pts))
  (setq c (cond ((and centroid (das-pt-in-poly-p centroid pts) pole
                      (>= (das-geom-edge-dist centroid pts) (* 0.8 (cadr pole))))
                 centroid)
                (pole (car pole))
                (T centroid)))
  ;; measure at rotation 0 and height h0
  (vl-catch-all-apply 'vla-put-Rotation (list obj 0.0))
  (vl-catch-all-apply 'vla-put-Height (list obj h0))
  (vl-catch-all-apply 'vla-Update (list obj))
  (setq ext (vl-catch-all-apply 'vla-GetBoundingBox (list obj 'mn 'mx)))
  (if (vl-catch-all-error-p ext)
    (setq w0 (* h0 9.0) t0 (* h0 1.4))          ; "S=1234.56 m2" estimate
    (setq mn (vlax-safearray->list mn) mx (vlax-safearray->list mx)
          w0 (- (car mx) (car mn)) t0 (- (cadr mx) (cadr mn))))
  ;; largest height that fits: along the parcel first, then horizontal
  (setq angs (if (equal ang 0.0 1e-9) (list 0.0) (list ang 0.0))
        hh h0 ok nil)
  (while (and (not ok) (>= hh (* hMin 0.999)))
    (foreach a angs
      (if (and (not ok)
               (das-label-fits-p c (* w0 (/ hh h0)) (* t0 (/ hh h0)) a pts (* 0.3 hh)))
        (setq ok a)))
    (if (not ok) (setq hh (* hh 0.85))))
  (if (not ok) (setq hh hMin ok 0.0))
  (vl-catch-all-apply 'vla-put-Height (list obj hh))
  (vl-catch-all-apply 'vla-put-Rotation (list obj ok))
  (vl-catch-all-apply 'vla-Update (list obj))
  ;; put the centre of the real box on the anchor, whatever the
  ;; attachment point or field evaluation did to it
  (setq ext (vl-catch-all-apply 'vla-GetBoundingBox (list obj 'mn 'mx)))
  (if (not (vl-catch-all-error-p ext))
    (progn
      (setq mn (vlax-safearray->list mn) mx (vlax-safearray->list mx))
      (vl-catch-all-apply 'vla-Move
        (list obj
              (vlax-3d-point (list (/ (+ (car mn) (car mx)) 2.0) (/ (+ (cadr mn) (cadr mx)) 2.0) 0.0))
              (vlax-3d-point (list (car c) (cadr c) 0.0))))))
  ent)

;; [FIX-16] A single summary label (unlike a dimension) has no edge it
;; has to align with, so orienting it along whichever edge happens to be
;; longest can leave it steep/near-vertical on plenty of real parcels -
;; harder to read, not easier. das-angle-is-steep-p measures how far an
;; angle sits from the nearest horizontal axis; das-render-all snaps the
;; label back to 0 when it would rather be reading like a dimension.
(defun das-angle-is-steep-p (ang / a)
  (setq a (rem (abs ang) pi))
  (if (> a (/ pi 2.0)) (setq a (- pi a)))
  (> a (/ pi 4.0)))

;; [NEW-1] -------------------------------------------------------------
;; Version 9.6 picked the vertex with the largest Y ("northernmost"),
;; with a min-X tie-break that a GPS survey never actually triggers.
;; On the Khormangah sample that chose a vertex 13.6 m EAST of the
;; parcel's west edge, i.e. the north-EAST corner.
;; das-geom-find-nw instead takes the vertex nearest the north-west
;; corner of the bounding box, which is what "شمال غرب" means.
;; ------------------------------------------------------------
;; [NEW-4] CORNER FILTER  -  هم‌راستا با سرویس utmdoc.ir
;; ------------------------------------------------------------
;; A vertex whose interior angle is close to 180 degrees is not a real
;; corner - it is a point that happens to sit along a straight side,
;; which is common with GPS collection. Such a point must never become
;; vertex number 1.
;;
;; IMPORTANT: these points are NOT removed from the parcel. They stay in
;; the drawing, the coordinate table and both Excel files, because they
;; may mark the boundary between two different neighbours. They are only
;; excluded from the list of candidates for the starting vertex.
;;
;; The threshold mirrors max_angle in cadastre.py on utmdoc.ir. Keep the
;; two numbers equal or the two products will number the same parcel
;; differently.
(defun das-geom-vertex-angle (pa pb pc / v1x v1y v2x v2y dot m1 m2 c)
  (setq v1x (- (car pa) (car pb))  v1y (- (cadr pa) (cadr pb))
        v2x (- (car pc) (car pb))  v2y (- (cadr pc) (cadr pb)))
  (setq dot (+ (* v1x v2x) (* v1y v2y))
        m1  (sqrt (+ (* v1x v1x) (* v1y v1y)))
        m2  (sqrt (+ (* v2x v2x) (* v2y v2y))))
  (if (or (zerop m1) (zerop m2))
    180.0
    (progn
      (setq c (/ dot (* m1 m2)))
      (cond ((> c  1.0) (setq c  1.0))
            ((< c -1.0) (setq c -1.0)))
      (/ (* 180.0 (atan (sqrt (- 1.0 (* c c))) c)) pi)
    )
  )
)

;; Indices of vertices that qualify as genuine corners.
;; If every vertex is nearly straight (a digitised circle, for example)
;; the whole list is returned rather than an empty one.
(defun das-geom-corner-indices (pts / i out)
  (setq i 0 out '())
  (foreach f (das-geom-corner-flags pts)
    (if f (setq out (cons i out)))
    (setq i (1+ i)))
  (reverse out)
)

;; [PERF-2] One flag per vertex (T = genuine corner), in vertex order,
;; so callers can walk it beside pts instead of looking up (nth i pts)
;; for every candidate index. Same all-straight fallback as above.
(defun das-geom-corner-flags (pts / out)
  (setq out (mapcar '(lambda (pa pb pc)
                       (<= (das-geom-vertex-angle pa pb pc) *das-corner-max-angle*))
                    (das-list-prev pts) pts (das-list-next pts)))
  (if (vl-some '(lambda (f) f) out) out (mapcar '(lambda (p) T) pts)))

(defun das-geom-all-indices (n / i out)
  (setq i 0 out '())
  (repeat n (setq out (cons i out) i (1+ i)))
  (reverse out)
)

;; NorthWest rule, metric variant (START_VERTEX_RULE = nw_metric).
;; ------------------------------------------------------------
;; Distance is measured in real metres from the north-west corner of the
;; bounding box. The web service also offers nw_normalized, which scales
;; X and Y by the box dimensions first; on a long narrow parcel the two
;; can pick different vertices. DAS implements the metric variant only,
;; so utmdoc.ir must be set to nw_metric to agree.
;;
;; Tie-break: if two candidates are within 1 cm of the corner distance,
;; the more western one wins. Without this, two runs on the same parcel
;; could differ purely from floating-point noise.
(defun das-geom-find-nw (pts / minx maxy corner flags i p d best bestp bestd)
  (setq minx (car (car pts)) maxy (cadr (car pts)))
  (foreach p pts
    (setq minx (min minx (car p)) maxy (max maxy (cadr p))))
  (setq corner (list minx maxy))
  (setq flags (das-geom-corner-flags pts) i 0 best nil bestp nil bestd nil)
  (foreach p pts
    (if (car flags)
      (progn
        (setq d (distance corner (list (car p) (cadr p))))
        (cond
          ((null bestd)             (setq bestd d best i bestp p))
          ((< d (- bestd 0.01))     (setq bestd d best i bestp p))
          ((and (<= d (+ bestd 0.01))
                (< (car p) (car bestp)))
                                    (setq bestd d best i bestp p))
        )))
    (setq flags (cdr flags) i (1+ i))
  )
  best
)

;; Northernmost vertex (9.6 behaviour). The corner filter applies here
;; too, so a point sitting in the middle of a straight northern side is
;; skipped in favour of a real corner.
(defun das-geom-find-north (pts / flags best bestp i p)
  (setq flags (das-geom-corner-flags pts) i 0 best nil bestp nil)
  (foreach p pts
    (if (and (car flags)
             (or (null bestp)
                 (> (cadr p) (cadr bestp))
                 (and (equal (cadr p) (cadr bestp) 1e-4) (< (car p) (car bestp)))))
      (setq bestp p best i))
    (setq flags (cdr flags) i (1+ i))
  )
  best
)

(defun das-geom-start-index (pts rule)
  (cond ((= rule "NorthWest")    (das-geom-find-nw pts))
        ((= rule "Northernmost") (das-geom-find-north pts))
        (T 0)))

(defun das-geom-rotate-list (lst idx / n)
  (setq n (length lst))
  (if (> n 0) (setq idx (rem idx n)))
  (append (das-list-drop lst idx) (das-list-take lst idx)))

;; Self-intersection test. A parcel whose boundary crosses itself has a
;; wrong area, and nothing downstream will notice.
(defun das--side (a b c)
  (- (* (- (car b) (car a)) (- (cadr c) (cadr a)))
     (* (- (cadr b) (cadr a)) (- (car c) (car a)))))

(defun das-geom-seg-cross-p (p1 p2 p3 p4 / d1 d2 d3 d4)
  (setq d1 (das--side p3 p4 p1) d2 (das--side p3 p4 p2)
        d3 (das--side p1 p2 p3) d4 (das--side p1 p2 p4))
  (and (< (* d1 d2) 0.0) (< (* d3 d4) 0.0)))

;; Edge i is (vertex i, vertex i+1). [PERF-2] Testing every pair of edges
;; is O(n^2) - 186 000 pairs on a 610-point boundary, most of the QC time.
;; Two edges can only cross when their bounding boxes overlap, so sort
;; the boxes by min X and compare each edge only with the edges that
;; start before it ends. Returns the same (i j) as the plain double loop:
;; the crossing pair with the smallest i, then the smallest j.
(defun das-geom-self-intersect-p (pts / n i recs a b rest lo hi best)
  (setq n (length pts) i -1 best nil)
  (setq recs
    (vl-sort
      (mapcar '(lambda (p q)
                 (list (min (car p) (car q)) (max (car p) (car q))
                       (min (cadr p) (cadr q)) (max (cadr p) (cadr q))
                       (setq i (1+ i)) p q))
              pts (das-list-next pts))
      '(lambda (a b) (< (car a) (car b)))))
  (while recs
    (setq a (car recs) rest (cdr recs))
    (while (and rest (<= (car (car rest)) (cadr a)))
      (setq b (car rest))
      (if (and (<= (nth 2 b) (nth 3 a)) (<= (nth 2 a) (nth 3 b)))
        (progn
          (setq lo (min (nth 4 a) (nth 4 b)) hi (max (nth 4 a) (nth 4 b)))
          (if (and (>= (- hi lo) 2)
                   (not (and (= lo 0) (= hi (1- n))))
                   (or (null best) (< lo (car best))
                       (and (= lo (car best)) (< hi (cadr best))))
                   (das-geom-seg-cross-p (nth 5 a) (nth 6 a) (nth 5 b) (nth 6 b)))
            (setq best (list lo hi)))))
      (setq rest (cdr rest)))
    (setq recs (cdr recs)))
  best)

;;; ============================================================
;;; MODULE 3C: GEOMETRY QC  [NEW-7, v10.3]
;;; ============================================================
;; Runs before the dialog (DAS/DASL) and on demand (DASQC). Each issue is
;; (key point). The table below gives every key a severity and its text:
;; ERR stops DAS unless the user confirms, WARN is only reported. Problem
;; points are circled in red on layer L-QC so they can be found on screen.
;; [NEW-15] The Farsi texts carry no Latin words: DCL text tiles and the
;; command line lay text out left-to-right, and a Latin word inside a
;; Farsi sentence splits it into pieces shown in the wrong order.
;; [FIX-20] touch replaces near: see das-qc-touch-points.
(setq *das-qc-text*
  '((few   ERR  "کمتر از سه رأس مجزا دارد"
               "fewer than 3 distinct vertices")
    (arc   ERR  "ضلع قوسی دارد؛ مساحت گواهی از مختصات رئوس حساب می‌شود و با شکل قوسی نمی‌خواند"
               "arc segment(s): the certificate area comes from the vertices and will not match the drawn shape")
    (cross ERR  "مرز قطعه خودش را قطع کرده است؛ مساحت معتبر نیست"
               "the boundary crosses itself: the area is not valid")
    (touch ERR  "مرز قطعه به خودش رسیده و قطعه دو تکه شده است؛ رأسی روی رأس یا ضلع دیگری افتاده است"
               "the boundary touches itself, so the parcel is pinched in two (a vertex lies on another vertex or side): area and numbering are not valid")
    (open  WARN "پلی‌لاین بسته نیست؛ بسته در نظر گرفته می‌شود"
               "polyline is not closed: DAS treats it as closed")
    (short WARN "ضلعی کوتاه‌تر از پنج سانتی‌متر دارد"
               "side shorter than 5 cm")
    (spike WARN "زاویهٔ داخلی تیزتر از پنج درجه دارد"
               "interior angle sharper than 5 degrees (spike)")
    (elev  WARN "ارتفاع رئوس صفر نیست؛ در گواهی فقط مختصات افقی می‌آید"
               "non-zero elevation (Z): only X and Y go on the certificate")
    (ocs   WARN "پلی‌لاین قرینه شده یا در دستگاه مختصات چرخیده رسم شده است؛ مختصات به دستگاه جهانی برگردانده شد"
               "polyline normal is not +Z (mirrored or tilted UCS): coordinates were converted to WCS")
    (utm   WARN "مختصات در محدودهٔ زون‌های ایران نیست؛ شاید نقشه در دستگاه مختصات دیگری است"
               "coordinates are outside the UTM range for Iran: the drawing may use another system")))

(defun das-qc-info (key) (assoc key *das-qc-text*))
(defun das-qc-sev  (key) (cadr (das-qc-info key)))
(defun das-qc-has-error-p (issues)
  (vl-some '(lambda (x) (eq (das-qc-sev (car x)) 'ERR)) issues))

;; Bulge of every raw vertex. LWPOLYLINE keeps them as DXF 42 on the
;; entity; an old-style POLYLINE keeps them on each VERTEX sub-entity.
(defun das-qc-bulges (ent / ed out e v)
  (setq ed (entget ent) out '())
  (if (= (cdr (assoc 0 ed)) "LWPOLYLINE")
    (foreach x ed (if (= (car x) 42) (setq out (cons (cdr x) out))))
    (progn
      (setq e (entnext ent))
      (while (and e (= (cdr (assoc 0 (setq v (entget e)))) "VERTEX"))
        (setq out (cons (cond ((cdr (assoc 42 v))) (0.0)) out)
              e   (entnext e)))))
  (reverse out))

;; [FIX-20] Boundary touching itself. das-geom-seg-cross-p only sees two
;; sides that pass THROUGH each other. Dropping a vertex onto another
;; vertex, onto another side, or running a side back along another one
;; leaves sides that merely touch: the parcel is pinched into two pieces
;; and nothing was reported but a "near vertices" warning (and for a
;; vertex on a side, nothing at all). A touch is a vertex closer than tol
;; to a side it does not belong to. Its own two sides are excluded, and
;; so are the sides one step further on (vertex j-1 or j+2 against side
;; j..j+1): a vertex there is only close because of a short side or a
;; spike, which have warnings of their own.
;; Sides are swept by X like das-geom-self-intersect-p, with their boxes
;; grown by tol. Returns the touching vertices, each once.
(defun das-pt-seg-dist (p a b / dx dy l2 u)
  (setq dx (- (car b) (car a)) dy (- (cadr b) (cadr a))
        l2 (+ (* dx dx) (* dy dy)))
  (if (zerop l2)
    (distance (list (car p) (cadr p)) (list (car a) (cadr a)))
    (progn
      (setq u (/ (+ (* (- (car p) (car a)) dx) (* (- (cadr p) (cadr a)) dy)) l2))
      (setq u (max 0.0 (min 1.0 u)))
      (distance (list (car p) (cadr p))
                (list (+ (car a) (* u dx)) (+ (cadr a) (* u dy)))))))

;; v = vertex index, vp = its point, s = side record, out = points so far
(defun das-qc-touch-check (v vp s n tol out / d)
  (setq d (rem (+ (- v (nth 4 s)) n n) n))                ; v - j, mod n
  (if (and (not (member d (list 0 1 2 (1- n))))
           (< (das-pt-seg-dist vp (nth 5 s) (nth 6 s)) tol)
           (not (vl-some '(lambda (q) (< (distance (list (car q) (cadr q)) (list (car vp) (cadr vp))) tol))
                         out)))
    (cons vp out)
    out))

(defun das-qc-touch-points (pts tol / n i recs a b rest out)
  (setq n (length pts) i -1 out '())
  (if (> n 4)
    (progn
      (setq recs
        (vl-sort
          (mapcar '(lambda (p q)
                     (list (- (min (car p) (car q)) tol) (+ (max (car p) (car q)) tol)
                           (- (min (cadr p) (cadr q)) tol) (+ (max (cadr p) (cadr q)) tol)
                           (setq i (1+ i)) p q))
                  pts (das-list-next pts))
          '(lambda (a b) (< (car a) (car b)))))
      (while recs
        (setq a (car recs) rest (cdr recs))
        (while (and rest (<= (car (car rest)) (cadr a)))
          (setq b (car rest))
          (if (and (<= (nth 2 b) (nth 3 a)) (<= (nth 2 a) (nth 3 b)))
            (setq out (das-qc-touch-check (nth 4 a)            (nth 5 a) b n tol out)
                  out (das-qc-touch-check (rem (1+ (nth 4 a)) n) (nth 6 a) b n tol out)
                  out (das-qc-touch-check (nth 4 b)            (nth 5 b) a n tol out)
                  out (das-qc-touch-check (rem (1+ (nth 4 b)) n) (nth 6 b) a n tol out)))
          (setq rest (cdr rest)))
        (setq recs (cdr recs)))))
  (reverse out))

;; The checks that need only the deduplicated points: crossing, touching,
;; then short sides and spikes per vertex. Returns the issues in order.
(defun das-qc-check-pts (pts / n hit issues pp p q)
  (setq n (length pts) issues '())
  ;; self-crossing (das-geom-self-intersect-p reports the first one)
  (if (setq hit (das-geom-self-intersect-p pts))
    (setq issues
      (cons (list 'cross
                  (cond ((inters (nth (car hit) pts) (nth (rem (1+ (car hit)) n) pts)
                                 (nth (cadr hit) pts) (nth (rem (1+ (cadr hit)) n) pts) T))
                        ((nth (car hit) pts))))
            issues)))
  ;; [FIX-20] sides touching (vertex on a vertex or side, 1 cm)
  (foreach p (das-qc-touch-points pts 0.01)
    (setq issues (cons (list 'touch p) issues)))
  ;; per vertex: short side to the next, spike
  (setq pp (das-list-prev pts) q (das-list-next pts))
  (foreach p pts
    (if (< (distance p (car q)) 0.05)
      (setq issues (cons (list 'short p) issues)))
    (if (< (das-geom-vertex-angle (car pp) p (car q)) 5.0)
      (setq issues (cons (list 'spike p) issues)))
    (setq pp (cdr pp) q (cdr q)))
  (reverse issues))

(defun das-qc-run (ent / obj ed raw pts n closed bl nb issues i r)
  (setq obj    (vlax-ename->vla-object ent)
        ed     (entget ent)
        raw    (das-geom-get-vertices ent)
        pts    (das-geom-dedupe raw 0.001)
        n      (length pts)
        closed (= (vlax-get obj 'Closed) -1)
        bl     (das-qc-bulges ent)
        issues '())
  (if (< n 3)
    (setq issues (list (list 'few (car raw))))
    (progn
      ;; arcs (the last vertex's bulge only counts on a closed polyline)
      (setq i 0 r raw nb (length bl))
      (foreach b bl
        (if (and (> (abs b) 1e-9) (or closed (< i (1- nb))))
          (setq issues (cons (list 'arc (car r)) issues)))
        (setq i (1+ i) r (cdr r)))
      ;; open polyline whose ends do not meet
      (if (and (not closed) (> (distance (car raw) (last raw)) 0.001))
        (setq issues (cons (list 'open (last raw)) issues)))
      ;; crossing, short sides, spikes, near-duplicate vertices
      (setq issues (append (reverse (das-qc-check-pts pts)) issues))
      ;; elevation
      (if (vl-some '(lambda (z) (> (abs z) 1e-6))
                   (if (= (cdr (assoc 0 ed)) "LWPOLYLINE")
                     (list (cond ((cdr (assoc 38 ed))) (0.0)))
                     (mapcar 'caddr raw)))
        (setq issues (cons (list 'elev (car pts)) issues)))
      ;; normal other than +Z
      (if (not (equal (cond ((cdr (assoc 210 ed))) ('(0.0 0.0 1.0))) '(0.0 0.0 1.0) 1e-9))
        (setq issues (cons (list 'ocs (car pts)) issues)))
      ;; UTM range for Iran (zones 38-41)
      (if (vl-some '(lambda (v) (or (< (car v) 160000.0) (> (car v) 840000.0)
                                    (< (cadr v) 2600000.0) (> (cadr v) 4500000.0)))
                   pts)
        (setq issues (cons (list 'utm (car pts)) issues)))))
  (reverse issues))

;; One line per issue type, with a count.
(defun das-qc-summary (issues / keys info cnt out en)
  (setq keys '() out "" en (= *das-lang* "English"))
  (foreach x issues (if (not (member (car x) keys)) (setq keys (cons (car x) keys))))
  (foreach k (reverse keys)
    (setq info (das-qc-info k) cnt 0)
    (foreach x issues (if (eq (car x) k) (setq cnt (1+ cnt))))
    (setq out (strcat out "\n"
                (if (eq (cadr info) 'ERR)
                  (if en "[ERROR] " "خطا: ")
                  (if en "[warning] " "هشدار: "))
                (if en (nth 3 info) (nth 2 info))
                (if (> cnt 1)
                  (if en (strcat " (" (itoa cnt) ")") (strcat "؛ " (das-fa-num cnt) " مورد"))
                  ""))))
  out)

(defun das-qc-mark (issues size / out)
  (das-utils-ensure-layer "L-QC" 1)
  (foreach x issues
    (if (cadr x)
      (setq out (cons (entmakex (list (cons 0 "CIRCLE") (cons 8 "L-QC") (cons 62 256)
                                      (cons 10 (cadr x)) (cons 40 size)))
                      out))))
  out)

(defun das-qc-clear-marks (/ ss i)
  (if (setq ss (ssget "_X" '((0 . "CIRCLE") (8 . "L-QC"))))
    (progn
      (setq i 0)
      (repeat (sslength ss) (entdel (ssname ss i)) (setq i (1+ i))))))

;;; ============================================================
;;; MODULE 3B: GEODESY
;;; ============================================================
;; [FIX-4] --------------------------------------------------------------
;; Snyder's inverse transverse-Mercator series is
;;   lat = phi1 - (N1 tanphi1 / R1)[ D^2/2 - (..)D^4/24 + (..)D^6/720 ]
;;   lon = lon0 + [ D - (..)D^3/6 + (..)D^5/120 ] / cos(phi1)
;; Version 9.6 wrote both tails as (- a b c), i.e. a-b-c, so the D^6 and
;; D^5 terms carried the wrong sign. The D^6 error is sub-millimetre and
;; harmless; the D^5 error reached roughly half a metre on the ground,
;; which matters on a certificate printed to 7 decimal degrees (~1 cm).
(defun das-math-utm-to-latlon (easting northing zone / a eccSq k0 lonOrigin
                               eccPrimeSq M mu e1 phi1Rad N1 tanPhi1 T1 C1 R1 D lat lon)
  (setq a          6378137.0
        eccSq      0.00669437999014
        k0         0.9996
        easting    (- easting 500000.0)
        lonOrigin  (+ (* (- zone 1) 6.0) -180.0 3.0)
        eccPrimeSq (/ eccSq (- 1.0 eccSq))
        M          (/ northing k0))
  ;; [FIX-9] Meridional arc series. The correct denominator is
  ;;     a[ 1 - e2/4 - 3*e2^2/64 - 5*e2^3/256 ]
  ;; Both 9.6 and 10.0 dropped the 3 and the 5. Because the resulting
  ;; error is constant (~5.7 m too far south) it never looks wrong on a
  ;; certificate - the latitude is plausible, just displaced. Verified
  ;; against pyproj: residual falls from 572 cm to 0.01 cm.
  (setq mu (/ M (* a (- 1.0 (/ eccSq 4.0)
                        (/ (* 3.0 eccSq eccSq) 64.0)
                        (/ (* 5.0 eccSq eccSq eccSq) 256.0)))))
  (setq e1 (/ (- 1.0 (sqrt (- 1.0 eccSq))) (+ 1.0 (sqrt (- 1.0 eccSq)))))
  (setq phi1Rad
    (+ mu
       (* (+ (* 3.0 (/ e1 2.0)) (* -27.0 (/ (* e1 e1 e1) 32.0))) (sin (* 2.0 mu)))
       (* (+ (* 21.0 (/ (* e1 e1) 16.0)) (* -55.0 (/ (* e1 e1 e1 e1) 32.0))) (sin (* 4.0 mu)))
       (* (* 151.0 (/ (* e1 e1 e1) 96.0)) (sin (* 6.0 mu)))
       (* (* 1097.0 (/ (* e1 e1 e1 e1) 512.0)) (sin (* 8.0 mu)))))
  (setq N1      (/ a (sqrt (- 1.0 (* eccSq (sin phi1Rad) (sin phi1Rad)))))
        tanPhi1 (/ (sin phi1Rad) (cos phi1Rad))
        T1      (* tanPhi1 tanPhi1)
        C1      (* eccPrimeSq (cos phi1Rad) (cos phi1Rad))
        R1      (/ (* a (- 1.0 eccSq))
                   (expt (- 1.0 (* eccSq (sin phi1Rad) (sin phi1Rad))) 1.5))
        D       (/ easting (* N1 k0)))
  (setq lat
    (- phi1Rad
       (* (/ (* N1 tanPhi1) R1)
          (+ (- (/ (* D D) 2.0)
                (/ (* (* D D D D)
                      (+ 5.0 (* 3.0 T1) (* 10.0 C1) (* -4.0 C1 C1) (* -9.0 eccPrimeSq)))
                   24.0))
             (/ (* (* D D D D D D)
                   (+ 61.0 (* 90.0 T1) (* 298.0 C1) (* 45.0 T1 T1)
                      (* -252.0 eccPrimeSq) (* -3.0 C1 C1)))
                720.0)))))
  (setq lon
    (+ (* (/ 180.0 pi)
          (/ (+ (- D (/ (* (* D D D) (+ 1.0 (* 2.0 T1) C1)) 6.0))
                (/ (* (* D D D D D)
                      (+ 5.0 (* -2.0 C1) (* 28.0 T1) (* -3.0 C1 C1)
                         (* 8.0 eccPrimeSq) (* 24.0 T1 T1)))
                   120.0))
             (cos phi1Rad)))
       lonOrigin))
  (list (* lat (/ 180.0 pi)) lon))

;;; ============================================================
;;; MODULE 4: EXCEL AUTOMATION
;;; ============================================================

;; [NEW-11] Find temp1.xlsx / temp2.xlsx (the Excel templates). Order:
;;   1. the folder remembered in the DAS_TEMPLATE_DIR setting,
;;   2. AutoCAD's support paths (file next to the lsp),
;;   3. a "templates" subfolder of a support path - keep the lsp and its
;;      templates\ folder together and add that folder to
;;      OPTIONS > Files > Support File Search Path,
;;   4. C:\DAS_Templates (the old fixed location).
;; If none has it, the user is asked once to point at the file and the
;; folder is remembered. NB: this must be a cond, not (or ...): in AutoLISP
;; (or "path" ...) returns T, not the string.
(defun das-find-template (fname / dir p)
  (setq dir (getenv "DAS_TEMPLATE_DIR") p nil)
  (cond
    ((and dir (/= dir "") (setq p (findfile (strcat dir "\\" fname)))) p)
    ((setq p (findfile fname)) p)
    ((setq p (findfile (strcat "templates\\" fname))) p)
    ((setq p (findfile (strcat "C:\\DAS_Templates\\" fname))) p)
    (T
     (setq p (getfiled (das-tr (strcat "محل فایل " fname " را نشان دهید")
                           (strcat "Locate the file " fname))
                     "" "xlsx" 0))
     (if p
       (progn (setenv "DAS_TEMPLATE_DIR" (vl-filename-directory p)) p)
       nil))))

(defun das-fix-vla (obj)
  (cond ((null obj) nil)
        ((= (type obj) 'variant) (das-fix-vla (vlax-variant-value obj)))
        (T obj)))

(defun das-excel-set-cell (ws row col val)
  (vlax-put-property
    (das-fix-vla (vlax-get-property (das-fix-vla ws) 'Cells)) 'Item row col val))

(defun das-excel-fmt (ws rng fmt)
  (vl-catch-all-apply 'vlax-put-property
    (list (das-fix-vla (vlax-get-property (das-fix-vla ws) 'Range rng))
          'NumberFormat fmt)))

(defun das-get-neighbor-text-for-excel (pIdx neighbors / res sIdx eIdx txt nb)
  (setq res "---")
  (foreach nb neighbors
    (setq sIdx (car nb) eIdx (cadr nb) txt (nth 2 nb))
    (if (or (and (<= sIdx pIdx) (< pIdx eIdx))
            (and (< eIdx sIdx) (or (>= pIdx sIdx) (< pIdx eIdx))))
      (setq res txt)))
  res)

;; Insert n rows at insRow, carrying the formatting of the row above.
(defun das-excel-grow (ws insRow n colSpan / k rowsItem rowObj srcRange destRange)
  (setq k 0)
  (repeat n
    (setq rowsItem (das-fix-vla (vlax-get-property ws 'Rows)))
    (setq rowObj   (das-fix-vla (vlax-get-property rowsItem 'Item insRow)))
    (vlax-invoke rowObj 'Insert)
    (setq srcRange  (das-fix-vla (vlax-get-property ws 'Range
                      (strcat "A" (itoa (1- insRow)) ":" colSpan (itoa (1- insRow))))))
    (setq destRange (das-fix-vla (vlax-get-property ws 'Range
                      (strcat "A" (itoa insRow) ":" colSpan (itoa insRow)))))
    (vlax-invoke srcRange 'Copy)
    (vlax-invoke destRange 'PasteSpecial -4122)   ; xlPasteFormats
    (setq insRow (1+ insRow) k (1+ k)))
  (princ))

;; Delete the template rows the parcel does not need. Without this, a
;; 10-point parcel left rows 21-24 of temp2 populated with the template's
;; own point numbers and live SQRT formulas pointing at empty cells.
(defun das-excel-shrink (ws firstRow lastRow / rangeObj)
  (if (>= lastRow firstRow)
    (vl-catch-all-apply
      '(lambda ()
        (setq rangeObj (das-fix-vla (vlax-get-property ws 'Range
                         (strcat (itoa firstRow) ":" (itoa lastRow)))))
        (vlax-invoke rangeObj 'Delete))
      '()))
  (princ))

;; [PERF-1] Writing one cell at a time (vlax-put-property ... 'Item row col)
;; means one COM round trip per cell. On a parcel with many vertices (a
;; GPS-collected boundary easily has a few hundred) that was the slowest
;; part of the whole export by far. A Range's Value2/Formula property
;; accepts a 2D SAFEARRAY and writes the whole block in a single COM call.
;; das-excel-array-1col / -2d build that array from a plain AutoLISP list;
;; das-excel-set-range / -formula assign it to a "A1:B9"-style range string.
(defun das-excel-array-1col (vals / n sa i v)
  (setq n (length vals)
        sa (vlax-make-safearray vlax-vbVariant (cons 0 (1- n)) (cons 0 0))
        i 0)
  (foreach v vals
    (vlax-safearray-put-element sa i 0 v)
    (setq i (1+ i)))
  (vlax-make-variant sa))

(defun das-excel-array-2d (rows / nr nc sa i j row v)
  (setq nr (length rows) nc (length (car rows))
        sa (vlax-make-safearray vlax-vbVariant (cons 0 (1- nr)) (cons 0 (1- nc)))
        i 0)
  (foreach row rows
    (setq j 0)
    (foreach v row
      (vlax-safearray-put-element sa i j v)
      (setq j (1+ j)))
    (setq i (1+ i)))
  (vlax-make-variant sa))

(defun das-excel-set-range (ws rangeStr variantArr)
  (vlax-put-property (das-fix-vla (vlax-get-property (das-fix-vla ws) 'Range rangeStr))
                      'Value2 variantArr))

(defun das-excel-set-range-formula (ws rangeStr variantArr)
  (vlax-put-property (das-fix-vla (vlax-get-property (das-fix-vla ws) 'Range rangeStr))
                      'Formula variantArr))

;; [FIX-19] SaveAs runs with DisplayAlerts off, so an existing output of
;; the same applicant used to be replaced without a word. Returns
;; (out1 out2), or nil when the user cancels. "New" picks the first
;; suffix _2, _3 ... for which neither file exists, so the pair stays
;; matched.
(defun das-excel-out-path (dir prefix name k)
  (strcat dir prefix name (if (> k 1) (strcat "_" (itoa k)) "") ".xlsx"))

(defun das-excel-pick-outputs (outDir safeName / k p1 p2 ans)
  (setq k 1
        p1 (das-excel-out-path outDir "Gvahi_Azla_" safeName k)
        p2 (das-excel-out-path outDir "ParcelMap_"  safeName k))
  (if (not (or (findfile p1) (findfile p2)))
    (list p1 p2)
    (progn
      (setq ans (das-ui-message
                  (das-tr "DAS - فایل اکسل تکراری" "DAS - Excel file exists")
                  (das-tr (strcat "فایل اکسل «" safeName "» در این پوشه از قبل وجود دارد:\n" (das-uchr 8206) outDir
                                  "\n\nفایل قبلی بازنویسی شود، فایل تازه با پسوند شماره ذخیره شود، یا صدور لغو شود؟")
                          (strcat "Excel output for \"" safeName "\" already exists in:\n" outDir
                                  "\n\nOverwrite it, save under a new name (_2, _3 ...), or cancel the export?"))
                  (list (list "Overwrite" (das-tr "بازنویسی" "Overwrite"))
                        (list "New"       (das-tr "فایل تازه" "New name"))
                        (list "Cancel"    (das-tr "انصراف" "Cancel")))
                  "New"))
      (cond
        ((= ans "Cancel") nil)
        ((= ans "Overwrite") (list p1 p2))
        (T
         (while (or (findfile p1) (findfile p2))
           (setq k  (1+ k)
                 p1 (das-excel-out-path outDir "Gvahi_Azla_" safeName k)
                 p2 (das-excel-out-path outDir "ParcelMap_"  safeName k)))
         (list p1 p2))))))

;; ---------------------------------------------------------------------
;; [FIX-5] The old code called vlax-get-or-create-object. If the user
;; already had Excel open, that returned THEIR session; the routine then
;; set Visible=false, DisplayAlerts=false and finally called Quit, which
;; closed their unsaved workbooks without a prompt. vlax-create-object
;; always starts a private instance.
;; [FIX-6] temp2 column E holds live =SQRT(...) formulas. 9.6 overwrote
;; them with rtos strings, and wrote B/C as text so the formulas would
;; have been fragile anyway. Now B/C are real numbers with an explicit
;; NumberFormat, column E is left alone, and the missing closing-segment
;; formula on the last row is supplied.
;; ---------------------------------------------------------------------
(defun das-export-excel-module (pts obj startNum neighbors areaVal /
                                centroid cx cy latlon lat lon excelApp wb1 ws1 wb2 ws2
                                row1 row2 temp1Path temp2Path
                                out1Path out2Path outDir workbooks numPts cap1 cap2
                                lastRow2 res safeName
                                nextPts closedFlag distList k pt nums adjList endRow1
                                xs ys nums2 adjList2 fList rowN outs)

  (setq centroid (das-geom-centroid pts)
        cx (car centroid) cy (cadr centroid)
        latlon (das-math-utm-to-latlon cx cy *das-zone*)
        lat (car latlon) lon (cadr latlon))

  (setq temp1Path (das-find-template "temp1.xlsx"))
  (setq temp2Path (if temp1Path (das-find-template "temp2.xlsx") nil))

  (setq outDir (getvar "DWGPREFIX"))
  (if (or (null outDir) (= outDir ""))
    (setq outDir (strcat (getenv "USERPROFILE") "\\Documents\\")))

  (setq safeName (das-safe-filename *das-applicant*))

  (cond
    ((not (and temp1Path temp2Path))
     ;; file and folder names on lines of their own (see das-fa-end)
     (alert (das-tr (strcat "فایل‌های الگوی اکسل پیدا نشدند.\n\n"
                            "این دو فایل را در پوشهٔ کنار فایل لیسپ بگذارید:\n"
                            "templates\\temp1.xlsx\ntemplates\\temp2.xlsx\n\n"
                            "یا در این پوشه:\nC:\\DAS_Templates")
                    (strcat "The Excel template files were not found.\n\n"
                            "Put these two files in the folder next to the LISP file:\n"
                            "templates\\temp1.xlsx\ntemplates\\temp2.xlsx\n\n"
                            "or in this folder:\nC:\\DAS_Templates"))))
    ((null (setq outs (das-excel-pick-outputs outDir safeName)))
     (princ (das-tr "\nصدور اکسل لغو شد." "\nExcel export cancelled.")))
    (T
      (setq out1Path (car outs) out2Path (cadr outs))
      (princ (das-tr "\nدر حال ساخت فایل‌های اکسل؛ چند لحظه صبر کنید…"
                     "\nCreating the Excel files, please wait..."))
      (setq numPts (length pts))

      (setq res (vl-catch-all-apply
        '(lambda ()
          (setq excelApp (das-fix-vla (vlax-create-object "Excel.Application")))
          (vlax-put-property excelApp 'Visible :vlax-false)
          (vlax-put-property excelApp 'DisplayAlerts :vlax-false)
          (setq workbooks (das-fix-vla (vlax-get-property excelApp 'Workbooks)))

          ;; ---------- temp1 : گواهی تعیین مختصات ----------
          ;; data rows 7..22 (16 bordered rows), note in row 23
          (setq wb1 (das-fix-vla (vlax-invoke workbooks 'Open temp1Path)))
          (setq ws1 (das-fix-vla (vlax-get-property wb1 'ActiveSheet)))
          (setq cap1 16 row1 7)
          (if (> numPts cap1) (das-excel-grow ws1 (+ row1 cap1) (- numPts cap1) "J"))

          ;; [PERF-1] One COM call per column instead of one per point+column:
          ;; nth on a list is itself O(n), so the old (nth i pts) loop was
          ;; quadratic in numPts on top of the per-cell Excel round trips.
          ;; Walking pts/nextPts in lockstep is a single linear pass.
          (setq nextPts (append (cdr pts) (list (car pts)))
                closedFlag (vlax-curve-isclosed obj)
                distList (mapcar 'distance pts nextPts))
          (if (not closedFlag)
            (setq distList (reverse (cons 0.0 (cdr (reverse distList))))))
          (setq nums '() k startNum)
          (foreach pt pts (setq nums (cons k nums) k (1+ k)))
          (setq nums (reverse nums))
          (setq adjList (mapcar '(lambda (n) (das-get-neighbor-text-for-excel n neighbors)) nums))
          (setq endRow1 (+ row1 numPts -1))
          (das-excel-set-range ws1 (strcat "A" (itoa row1) ":B" (itoa endRow1))
                                (das-excel-array-2d (mapcar 'list nums distList)))
          (das-excel-set-range ws1 (strcat "D" (itoa row1) ":D" (itoa endRow1))
                                (das-excel-array-1col adjList))

          (das-excel-set-cell ws1 3 4 lon)
          (das-excel-set-cell ws1 4 4 lat)
          (das-excel-fmt ws1 "D3:D4" "0.0000000")
          ;; drop the unused template rows (7+numPts .. 22)
          (if (< numPts cap1) (das-excel-shrink ws1 (+ 7 numPts) (+ 7 cap1 -1)))
          (vlax-put-property excelApp 'CutCopyMode 0)
          (vlax-invoke wb1 'SaveAs out1Path)
          (vlax-invoke wb1 'Close :vlax-false)
          (setq wb1 nil)

          ;; ---------- temp2 : ParcelMap ----------
          ;; data rows 11..24 (14 rows), E holds =SQRT(...) formulas
          (setq wb2 (das-fix-vla (vlax-invoke workbooks 'Open temp2Path)))
          (setq ws2 (das-fix-vla (vlax-get-property wb2 'ActiveSheet)))
          (setq cap2 14 row2 11)
          (if (> numPts cap2) (das-excel-grow ws2 (+ row2 cap2) (- numPts cap2) "E"))

          (das-excel-set-cell ws2 2 3 *das-applicant*)
          (das-excel-set-cell ws2 3 3 *das-nat-id*)
          (das-excel-set-cell ws2 5 3 areaVal)
          (das-excel-set-cell ws2 6 3 *das-b-area*)
          (das-excel-set-cell ws2 8 3 *das-zone*)
          (das-excel-set-cell ws2 9 3 numPts)

          (setq lastRow2 (+ 11 numPts -1))
          (das-excel-fmt ws2 (strcat "B11:C" (itoa lastRow2)) "0.0000")
          (das-excel-fmt ws2 (strcat "E11:E" (itoa lastRow2)) "0.00")

          (setq nums2 '() k startNum)
          (foreach pt pts (setq nums2 (cons k nums2) k (1+ k)))
          (setq nums2 (reverse nums2))
          (setq adjList2 (mapcar '(lambda (n) (das-get-neighbor-text-for-excel n neighbors)) nums2))
          (setq xs (mapcar 'car pts) ys (mapcar 'cadr pts))
          (das-excel-set-range ws2 (strcat "A" (itoa row2) ":D" (itoa lastRow2))
                                (das-excel-array-2d (mapcar 'list nums2 xs ys adjList2)))

          ;; column E: keep the template's own formula where it is still
          ;; valid (numPts <= 14, every row but the last). Otherwise every
          ;; row needs a fresh formula - write them all in one call instead
          ;; of one SQRT formula per row.
          (if (> numPts cap2)
            (progn
              (setq fList '() rowN row2)
              (while (< rowN lastRow2)
                (setq fList (cons (strcat "=SQRT((B" (itoa rowN) "-B" (itoa (1+ rowN)) ")^2+(C"
                                          (itoa rowN) "-C" (itoa (1+ rowN)) ")^2)") fList)
                      rowN (1+ rowN)))
              (setq fList (cons (strcat "=SQRT((B" (itoa lastRow2) "-B11)^2+(C" (itoa lastRow2) "-C11)^2)") fList))
              (das-excel-set-range-formula ws2 (strcat "E11:E" (itoa lastRow2))
                                            (das-excel-array-1col (reverse fList))))
            (das-excel-set-cell ws2 lastRow2 5
              (strcat "=SQRT((B" (itoa lastRow2) "-B11)^2+(C" (itoa lastRow2) "-C11)^2)")))

          (das-excel-fmt ws2 "C5:C6" "0.00")
          ;; drop the unused template rows (11+numPts .. 24)
          (if (< numPts cap2) (das-excel-shrink ws2 (+ 11 numPts) (+ 11 cap2 -1)))
          (vlax-put-property excelApp 'CutCopyMode 0)
          (vlax-invoke wb2 'SaveAs out2Path)
          (vlax-invoke wb2 'Close :vlax-false)
          (setq wb2 nil)
          T)
        '()))

      ;; always tear the private Excel instance down, success or not
      (if excelApp
        (progn
          (vl-catch-all-apply '(lambda () (if wb1 (vlax-invoke wb1 'Close :vlax-false))) '())
          (vl-catch-all-apply '(lambda () (if wb2 (vlax-invoke wb2 'Close :vlax-false))) '())
          (vl-catch-all-apply '(lambda () (vlax-invoke excelApp 'Quit)) '())
          (vl-catch-all-apply '(lambda () (vlax-release-object excelApp)) '())
          (setq excelApp nil)
          (gc)))

      (if (vl-catch-all-error-p res)
        ;; wb1 is saved before wb2 is opened, so a failure in temp2 leaves
        ;; the first file on disk; "nothing was saved" was not true
        (alert (das-tr (strcat "ساخت فایل‌های اکسل کامل نشد:\n"
                               (vl-catch-all-error-message res)
                               "\n\nممکن است فقط فایل اول ذخیره شده باشد؛ صدور را دوباره انجام دهید.")
                       (strcat "The Excel files were not completed:\n"
                               (vl-catch-all-error-message res)
                               "\n\nOnly the first file may have been saved; run the export again.")))
        (alert (das-tr (strcat "فایل‌های اکسل ذخیره شدند.\nپوشه:\n" outDir
                               "\n\n" (vl-filename-base out1Path) ".xlsx"
                               "\n" (vl-filename-base out2Path) ".xlsx")
                       (strcat "The Excel files were saved.\nFolder:\n" outDir
                               "\n\n" (vl-filename-base out1Path) ".xlsx"
                               "\n" (vl-filename-base out2Path) ".xlsx"))))))
  (princ))

;;; ============================================================
;;; MODULE 5: DCL CALLBACKS
;;; ============================================================

;; [NEW-6] Shared by the dialog's "+ Add" (the edge-walk caller was removed in 10.2.2), so the
;; two entry paths can never disagree on wording for the same input.
;; Returns (list dataEntry displayStr) where dataEntry is what
;; *das-temp-nb-data* stores (si ei composedText) and displayStr is what
;; the "lb_neighbors" list box shows.
(defun das-nb-compose (si ei n_type n_name / composed display)
;; [FIX-26] The sentence that is drawn and exported, and the list-box line, are
  ;; Farsi whatever the interface language is: the tool is for Iran and
  ;; English is only the look of the dialog. "to" is never used. Only where
  ;; Farsi text cannot exist (see das-farsi-ok-p) a plain " - " joins the two.
  (setq composed (if (das-unicode-p)
                   (strcat n_type "ی است به " n_name)
                   (strcat n_type " - " n_name)))
  (setq display (if (das-farsi-ok-p)
                  (strcat "P" (itoa si) " -> P" (itoa ei) ": " n_type " به " n_name)
                  (strcat "P" (itoa si) " -> P" (itoa ei) ": " n_type " - " n_name)))
  (list (list si ei composed) display))

;; [NEW-19] Free text instead of the "type + name" pattern. It is stored
;; and shown exactly as typed, so it replaces the pattern's sentence on the
;; drawing, in the list box and in both Excel files.
(defun das-nb-compose-custom (si ei text)
  (list (list si ei text)
        (strcat "P" (itoa si) " -> P" (itoa ei) ": " text)))

(defun das-ui-add-neighbor ( / n_start n_end n_type n_name n_custom si ei nmax pair)
  (setq n_start (get_tile "eb_nb_start")
        n_end   (get_tile "eb_nb_end")
        n_type  (get_tile "eb_nb_type")
        n_name  (get_tile "eb_nb_name")
        n_custom (vl-string-trim " \t" (get_tile "eb_nb_custom")))   ; [NEW-19]
  (setq si (atoi n_start) ei (atoi n_end))
  (setq nmax (+ *das-last-start* (1- *das-nvert*)))
  (cond
    ((or (= n_start "") (= n_end "") (and (= n_name "") (= n_custom "")))
     (set_tile "error" (if (= *das-lang* "Farsi")
                         "شمارهٔ دو نقطه و نام مجاور (یا متن دلخواه) را وارد کنید."
                         "Enter both point numbers and a neighbour name (or custom text).")))
    ;; [FIX-3] 9.6 never checked the range, so a typo produced
    ;; (nth 99 vertices) -> nil -> crash inside das-draw-range-neighbors
    ((or (< si *das-last-start*) (> si nmax) (< ei *das-last-start*) (> ei nmax))
     (set_tile "error" (das-tr (strcat "شمارهٔ نقطه باید بین " (das-fa-num *das-last-start*)
                                       " و " (das-fa-num nmax) " باشد.")
                               (strcat "Point number must be between " (itoa *das-last-start*)
                                       " - " (itoa nmax)))))
    ((= si ei)
     (set_tile "error" (if (= *das-lang* "Farsi")
                         "نقطهٔ شروع و پایان نباید یکی باشند."
                         "Start and end point must differ.")))
    (T
     (set_tile "error" "")
     (setq pair (if (/= n_custom "")
                  (das-nb-compose-custom si ei n_custom)
                  (das-nb-compose si ei n_type n_name)))
     (setq *das-temp-nb-data*    (append *das-temp-nb-data*    (list (car pair))))
     (setq *das-temp-nb-strings* (append *das-temp-nb-strings* (list (cadr pair))))
     (start_list "lb_neighbors") (mapcar 'add_list *das-temp-nb-strings*) (end_list)
     (set_tile "eb_nb_start" "") (set_tile "eb_nb_end" "") (set_tile "eb_nb_name" "")
     (set_tile "eb_nb_custom" "")))
  (princ))

(defun das-ui-del-neighbor ( / idx)
  (setq idx (atoi (get_tile "lb_neighbors")))
  (if (and *das-temp-nb-strings* (< idx (length *das-temp-nb-strings*)))
    (progn
      (setq *das-temp-nb-strings* (das-utils-remove-at idx *das-temp-nb-strings*))
      (setq *das-temp-nb-data*    (das-utils-remove-at idx *das-temp-nb-data*))
      (start_list "lb_neighbors") (mapcar 'add_list *das-temp-nb-strings*) (end_list)))
  (princ))

;; [FIX-3] Every value is clamped. 9.6 passed atoi/atof output straight
;; through: an empty column-count field became 0 and then divided by
;; zero in das-draw-table; a zero text scale produced nil entities.
;; [NEW-12] On OK (not Preview) a national ID that fails the check is
;; refused once; pressing OK again with the same value keeps it.
(defun das-ui-dcl-accept-callback (status_code / nid)
  (setq nid (das-digits-ascii (get_tile "eb_nid")))
  (if (and (= status_code 1) (/= nid "")
           (not (das-national-id-valid-p nid))
           (/= nid *das-nid-warned*))
    (progn
      (setq *das-nid-warned* nid)
      (set_tile "error" (das-tr (das-fa-end "کد ملی معتبر نیست؛ رقم کنترل آن نمی‌خواند. اگر از درستی‌اش مطمئنید، دوباره بزنید:" "OK")
                                "National ID fails the check digit. Press OK again to keep it anyway."))
      (mode_tile "eb_nid" 2))
    (das-ui-dcl-accept-apply status_code nid)))

(defun das-ui-dcl-accept-apply (status_code nid)
  (setq *das-last-scale*     (das-clamp     (atof (get_tile "eb_scale"))      0.01 1000.0)
        *das-last-start*      (das-clamp-int (atoi (get_tile "eb_start"))         0 99999)
        *das-last-offset*     (das-clamp     (atof (get_tile "eb_offset"))      0.0 10000.0)
        *das-last-prec-c*     (das-clamp-int (atoi (get_tile "eb_prec_c"))        0 8)
        *das-last-prec-l*     (das-clamp-int (atoi (get_tile "eb_prec_l"))        0 8)
        *das-last-prec-a*     (das-clamp-int (atoi (get_tile "eb_prec_a"))        0 8)
        *das-last-mksize*     (das-clamp     (atof (get_tile "eb_mksize"))    0.001 100.0)
        *das-last-mask*       (if (= (get_tile "tg_mask") "1") "Yes" "No")
        *das-last-withlen*    (= (get_tile "tg_withlen") "1")
        *das-last-small-th*   (das-clamp     (atof (get_tile "eb_small_th"))    0.0 10000.0)
        *das-last-small-fac*  (das-clamp     (atof (get_tile "eb_small_fac"))   0.1 5.0)
        *das-last-col-count*  (das-clamp-int (atoi (get_tile "eb_col_count"))     1 10)
        *das-last-mkstyle*    (nth (atoi (get_tile "pl_mkstyle"))
                                '("None" "Circle" "Square" "Triangle" "Xmark"))
        *das-start-rule*      (nth (atoi (get_tile "pl_startrule"))
                                '("NorthWest" "Northernmost" "AsDrawn"))
        *das-last-nb-scale*   (das-clamp     (atof (get_tile "eb_nb_scale"))   0.01 1000.0)
        *das-last-nb-offset*  (das-clamp     (atof (get_tile "eb_nb_offset"))   0.0 10000.0)
        *das-applicant*       (vl-string-trim " " (get_tile "eb_app"))
        *das-nat-id*          nid
        *das-b-area*          (das-clamp     (atof (get_tile "eb_barea"))       0.0 1e9)
        *das-zone*            (das-clamp-int (atoi (get_tile "eb_zone"))          1 60))
  (done_dialog status_code))

;;; ============================================================
;;; MODULE 6: GUI DIALOG
;;; ============================================================
(defun das-ui-dcl-show (defaultMarkerSize / dcl_file file_id dcl_id status styles_list match_idx
                         lbl_main lbl_box1 lbl_scale lbl_start lbl_offset lbl_box2 lbl_precc
                         lbl_precl lbl_preca lbl_box3 lbl_mkstyle lbl_mksize lbl_mask lbl_box4
                         lbl_smallth lbl_smallfac lbl_colcount lbl_box5 lbl_nbstart lbl_nbend
                         lbl_nbtype lbl_nbname lbl_nbcustom lbl_nbscale lbl_nboffset lbl_btnadd lbl_btndel
                         lbl_btnprev lbl_box6 lbl_app lbl_nid lbl_barea lbl_zone
                         lbl_withlen lbl_startrule)

  (if (= *das-lang* "Farsi")
    (setq lbl_main      (strcat "DAS Tools V" *das-version* " - مدیریت هوشمند پلاک ثبتی")
          lbl_box1      "تنظیمات متون و ابعاد"
          lbl_scale     "مقیاس متن (Text Scale):"
          lbl_start     "شماره شروع رئوس:"
          lbl_offset    "فاصلهٔ عدد طول از ضلع:"
          lbl_startrule "قاعده نقطه شروع:"
          lbl_box2      "دقت اعشار (Precision)"
          lbl_precc     "مختصات (Coordinate):"
          lbl_precl     "طول اضلاع (Dimension):"
          lbl_preca     "مساحت پلاک (Area):"
          lbl_box3      "مارکرها و کارتوگرافی"
          lbl_mkstyle   "نوع مارکر گوشه:"
          lbl_mksize    "سایز هندسی مارکر:"
          lbl_mask      "اعمال کادر سفید (TextMask)"
          lbl_box4      "اضلاع کوتاه و جدول"
          lbl_smallth   "آستانه طول کوچک (0=خاموش):"
          lbl_smallfac  "ضریب کوچک‌نمایی متن:"
          lbl_colcount  "تعداد ستون‌های جدول مختصات:"
          lbl_withlen   "درج ستون طول اضلاع در جدول"
          lbl_box5      "مدیریت و درج مجاورین ثبتی"
          lbl_nbstart   "از پوینت:"
          lbl_nbend     "تا پوینت:"
          lbl_nbtype    "نوع حد (مثلاً دیواری):"
          lbl_nbname    "نام مجاور (مثلاً کوچه/پلاک):"
          lbl_nbcustom  "متن دلخواه (جای الگو را می‌گیرد):"
          lbl_nbscale   "سایز متن مجاور:"
          lbl_nboffset  "آفست حد مجاور:"
          lbl_btnadd    "+ افزودن به لیست"
          lbl_btndel    "X حذف انتخابی"
          lbl_btnprev   "پیش نمایش زنده روی نقشه"
          lbl_box6      "مشخصات پروژه جهت خروجی کاداستر"
          lbl_app       "نام متقاضی:"
          lbl_nid       "شماره ملی:"
          lbl_barea     "مساحت بنا:"
          lbl_zone      "زون UTM:")
    (setq lbl_main      (strcat "DAS Tools V" *das-version* " - Smart Cadastral Management")
          lbl_box1      "Text and Dimension Settings"
          lbl_scale     "Text Scale:"
          lbl_start     "Start Number:"
          lbl_offset    "Length Text Offset:"
          lbl_startrule "Start Vertex Rule:"
          lbl_box2      "Precision Settings"
          lbl_precc     "Coordinate Decimals:"
          lbl_precl     "Dimension Decimals:"
          lbl_preca     "Area Decimals:"
          lbl_box3      "Markers and Cartography"
          lbl_mkstyle   "Corner Marker Type:"
          lbl_mksize    "Marker Size:"
          lbl_mask      "Enable Text Mask (Wipeout)"
          lbl_box4      "Small Edges and Table"
          lbl_smallth   "Small Edge Threshold (0=Off):"
          lbl_smallfac  "Small Text Scale Factor:"
          lbl_colcount  "Table Column Count:"
          lbl_withlen   "Include edge Length column"
          lbl_box5      "Cadastral Neighbors Management"
          lbl_nbstart   "From Point:"
          lbl_nbend     "To Point:"
          lbl_nbtype    "Boundary Type (e.g. Wall):"
          lbl_nbname    "Neighbor Name (e.g. Street):"
          lbl_nbcustom  "Custom Text (replaces the pattern):"
          lbl_nbscale   "Neighbor Text Size:"
          lbl_nboffset  "Neighbor Offset:"
          lbl_btnadd    "+ Add to List"
          lbl_btndel    "X Delete Selected"
          lbl_btnprev   "Live Preview on Map"
          lbl_box6      "Project Details for Excel Export"
          lbl_app       "Applicant Name:"
          lbl_nid       "National ID:"
          lbl_barea     "Bldg Area:"
          lbl_zone      "UTM Zone:"))

  (setq dcl_file (vl-filename-mktemp "das_tools" nil ".dcl"))
  (setq file_id (open dcl_file "w"))
  (write-line "das_gui_dialog : dialog {" file_id)
  (write-line (strcat "  label = \"" lbl_main "\";") file_id)
  (write-line "  : row {" file_id)
  (write-line "    : column {" file_id)
  (write-line (strcat "      : boxed_column { label = \"" lbl_box1 "\";") file_id)
  (write-line (strcat "        : edit_box { key = \"eb_scale\"; label = \"" lbl_scale "\"; edit_width = 10; }") file_id)
  (write-line (strcat "        : edit_box { key = \"eb_start\"; label = \"" lbl_start "\"; edit_width = 10; }") file_id)
  (write-line (strcat "        : edit_box { key = \"eb_offset\"; label = \"" lbl_offset "\"; edit_width = 10; }") file_id)
  (write-line (strcat "        : popup_list { key = \"pl_startrule\"; label = \"" lbl_startrule "\"; width = 15; }") file_id)
  (write-line "      }" file_id)
  (write-line (strcat "      : boxed_column { label = \"" lbl_box2 "\";") file_id)
  (write-line (strcat "        : edit_box { key = \"eb_prec_c\"; label = \"" lbl_precc "\"; edit_width = 10; }") file_id)
  (write-line (strcat "        : edit_box { key = \"eb_prec_l\"; label = \"" lbl_precl "\"; edit_width = 10; }") file_id)
  (write-line (strcat "        : edit_box { key = \"eb_prec_a\"; label = \"" lbl_preca "\"; edit_width = 10; }") file_id)
  (write-line "      }" file_id)
  (write-line (strcat "      : boxed_column { label = \"" lbl_box3 "\";") file_id)
  (write-line (strcat "        : popup_list { key = \"pl_mkstyle\"; label = \"" lbl_mkstyle "\"; width = 15; }") file_id)
  (write-line (strcat "        : edit_box { key = \"eb_mksize\"; label = \"" lbl_mksize "\"; edit_width = 10; }") file_id)
  (write-line (strcat "        : toggle { key = \"tg_mask\"; label = \"" lbl_mask "\"; }") file_id)
  (write-line "      }" file_id)
  (write-line (strcat "      : boxed_column { label = \"" lbl_box4 "\";") file_id)
  (write-line (strcat "        : edit_box { key = \"eb_small_th\"; label = \"" lbl_smallth "\"; edit_width = 10; }") file_id)
  (write-line (strcat "        : edit_box { key = \"eb_small_fac\"; label = \"" lbl_smallfac "\"; edit_width = 10; }") file_id)
  (write-line (strcat "        : edit_box { key = \"eb_col_count\"; label = \"" lbl_colcount "\"; edit_width = 10; }") file_id)
  (write-line (strcat "        : toggle { key = \"tg_withlen\"; label = \"" lbl_withlen "\"; }") file_id)
  (write-line "      }" file_id)
  (write-line "    }" file_id)
  (write-line "    : column {" file_id)
  (write-line (strcat "      : boxed_column { label = \"" lbl_box5 "\";") file_id)
  (write-line "        : row {" file_id)
  (write-line (strcat "          : edit_box { key = \"eb_nb_start\"; label = \"" lbl_nbstart "\"; edit_width = 5; }") file_id)
  (write-line (strcat "          : edit_box { key = \"eb_nb_end\"; label = \"" lbl_nbend "\"; edit_width = 5; }") file_id)
  (write-line "        }" file_id)
  (write-line (strcat "        : edit_box { key = \"eb_nb_type\"; label = \"" lbl_nbtype "\"; edit_width = 22; }") file_id)
  (write-line (strcat "        : edit_box { key = \"eb_nb_name\"; label = \"" lbl_nbname "\"; edit_width = 22; }") file_id)
  (write-line (strcat "        : edit_box { key = \"eb_nb_custom\"; label = \"" lbl_nbcustom "\"; edit_width = 22; }") file_id)
  (write-line "        : row {" file_id)
  (write-line (strcat "          : edit_box { key = \"eb_nb_scale\"; label = \"" lbl_nbscale "\"; edit_width = 6; }") file_id)
  (write-line (strcat "          : edit_box { key = \"eb_nb_offset\"; label = \"" lbl_nboffset "\"; edit_width = 6; }") file_id)
  (write-line "        }" file_id)
  (write-line "        spacer;" file_id)
  (write-line "        : row {" file_id)
  (write-line (strcat "          : button { key = \"btn_nb_add\"; label = \"" lbl_btnadd "\"; }") file_id)
  (write-line (strcat "          : button { key = \"btn_nb_del\"; label = \"" lbl_btndel "\"; }") file_id)
  (write-line "        }" file_id)
  (write-line "        spacer;" file_id)
  (write-line "        : list_box { key = \"lb_neighbors\"; width = 25; height = 9; }" file_id)
  (write-line "      }" file_id)
  (write-line (strcat "      : boxed_column { label = \"" lbl_box6 "\";") file_id)
  (write-line "        : row {" file_id)
  (write-line (strcat "          : edit_box { key = \"eb_app\"; label = \"" lbl_app "\"; edit_width = 15; }") file_id)
  (write-line (strcat "          : edit_box { key = \"eb_nid\"; label = \"" lbl_nid "\"; edit_width = 15; }") file_id)
  (write-line "        }" file_id)
  (write-line "        : row {" file_id)
  (write-line (strcat "          : edit_box { key = \"eb_barea\"; label = \"" lbl_barea "\"; edit_width = 15; }") file_id)
  (write-line (strcat "          : edit_box { key = \"eb_zone\"; label = \"" lbl_zone "\"; edit_width = 15; }") file_id)
  (write-line "        }" file_id)
  (write-line "      }" file_id)
  (write-line "    }" file_id)
  (write-line "  }" file_id)
  (write-line "  errtile;" file_id)
  (write-line "  spacer;" file_id)
  (write-line "  : row {" file_id)
  (write-line (strcat "    : button { key = \"btn_lang\"; label = \""
                      (cond ((not (das-farsi-ok-p)) "Farsi UI unavailable (see guide)")   ; [FIX-25]
                            ((= *das-lang* "Farsi") "English")
                            (T "فارسی"))
                      "\";" (if (das-farsi-ok-p) "" " is_enabled = false;") " }") file_id)
  (write-line "    : spacer { width = 2; }" file_id)
  (write-line (strcat "    : button { key = \"btn_preview\"; label = \"" lbl_btnprev "\"; }") file_id)
  (write-line "    : spacer { width = 4; }" file_id)
  (write-line "    : button { key = \"accept\"; label = \"OK\"; is_default = true; }" file_id)
  (write-line "    : button { key = \"cancel\"; label = \"Cancel\"; is_cancel = true; }" file_id)
  (write-line "  }" file_id)
  (write-line "}" file_id)
  (close file_id)

  (setq dcl_id (load_dialog dcl_file))
  (if (not (new_dialog "das_gui_dialog" dcl_id))
    (progn (vl-file-delete dcl_file) (setq status 0))
    (progn
      (set_tile "eb_scale"      (rtos *das-last-scale* 2 2))
      (set_tile "eb_start"      (itoa *das-last-start*))
      (set_tile "eb_offset"     (rtos (if *das-last-offset* *das-last-offset* (* *das-last-scale* 2.0)) 2 2))
      (set_tile "eb_prec_c"     (itoa *das-last-prec-c*))
      (set_tile "eb_prec_l"     (itoa *das-last-prec-l*))
      (set_tile "eb_prec_a"     (itoa *das-last-prec-a*))
      (set_tile "eb_mksize"     (rtos (if *das-last-mksize* *das-last-mksize* defaultMarkerSize) 2 2))
      (set_tile "tg_mask"       (if (= *das-last-mask* "Yes") "1" "0"))
      (set_tile "tg_withlen"    (if *das-last-withlen* "1" "0"))
      (set_tile "eb_small_th"   (rtos *das-last-small-th* 2 2))
      (set_tile "eb_small_fac"  (rtos *das-last-small-fac* 2 2))
      (set_tile "eb_col_count"  (itoa *das-last-col-count*))
      (set_tile "eb_nb_scale"   (rtos *das-last-nb-scale* 2 2))
      (set_tile "eb_nb_offset"  (rtos *das-last-nb-offset* 2 2))
      (set_tile "eb_app"        *das-applicant*)
      (set_tile "eb_nid"        *das-nat-id*)
      (set_tile "eb_barea"      (rtos *das-b-area* 2 2))
      (set_tile "eb_zone"       (itoa *das-zone*))

      (setq styles_list '("None" "Circle" "Square" "Triangle" "Xmark"))
      (start_list "pl_mkstyle") (mapcar 'add_list styles_list) (end_list)
      (setq match_idx (vl-position *das-last-mkstyle* styles_list))
      (set_tile "pl_mkstyle" (itoa (if match_idx match_idx 1)))

      (start_list "pl_startrule")
      (mapcar 'add_list (if (= *das-lang* "Farsi")
                          '("شمال غربی (استاندارد)" "شمالی‌ترین" "بدون تغییر")
                          '("North-West (standard)" "Northernmost" "As drawn")))
      (end_list)
      (setq match_idx (vl-position *das-start-rule* '("NorthWest" "Northernmost" "AsDrawn")))
      (set_tile "pl_startrule" (itoa (if match_idx match_idx 0)))

      (start_list "lb_neighbors") (mapcar 'add_list *das-temp-nb-strings*) (end_list)

      (action_tile "btn_nb_add"  "(das-ui-add-neighbor)")
      (action_tile "btn_nb_del"  "(das-ui-del-neighbor)")
      (action_tile "btn_preview" "(das-ui-dcl-accept-callback 2)")
      ;; [NEW-17] keep what was typed, close, reopen in the other language
      (action_tile "btn_lang"    "(das-ui-dcl-accept-apply 3 (das-digits-ascii (get_tile \"eb_nid\")))")
      (action_tile "accept"      "(das-ui-dcl-accept-callback 1)")
      (action_tile "cancel"      "(done_dialog 0)")

      (setq status (start_dialog))
      (unload_dialog dcl_id)
      (vl-file-delete dcl_file)))
  status)

;;; ============================================================
;;; MODULE 6B: MESSAGE WINDOW AND NOTES  [NEW-15, v10.6]
;;; ============================================================
;; Command-line messages scroll away and are easy to miss. das-ui-message
;; shows a small DCL window: title, the text one line per row, and a row
;; of buttons. buttons = ((key label) ...); the one whose key is dflt is
;; the default (Enter), the last one answers Esc. Returns the key pressed.
;; alert cannot do this: it has an OK button only. If the window cannot be
;; shown (no DCL, e.g. accoreconsole) it falls back to alert and returns
;; dflt, so dflt must be the safe answer.
(defun das-dcl-esc (s / out i c)
  (setq out "" i 0)
  (repeat (strlen s)
    (setq i (1+ i) c (substr s i 1))
    (setq out (strcat out (cond ((= c "\\") "\\\\") ((= c "\"") "\\\"") ((= c "\t") "    ") (T c)))))
  out)

(defun das-split-lines (s / out cur i c)
  (setq out '() cur "" i 0)
  (repeat (strlen s)
    (setq i (1+ i) c (substr s i 1))
    (if (= c "\n") (setq out (cons cur out) cur "") (setq cur (strcat cur c))))
  (setq out (reverse (cons cur out)))
  (while (and out (= (vl-string-trim " " (car out)) "")) (setq out (cdr out)))
  out)

;; A DCL text tile (like the command line) lays every line out with a
;; LEFT-to-right base direction; a mark at the start does not change
;; that, and the embedding controls that would (U+202B..) do not exist in
;; code page 1256, which the DCL file is written in. So the messages are
;; worded for it: no Latin word inside a Farsi sentence, at most one at
;; its very end (das-fa-end), digits only after Farsi text. Each Farsi
;; line is right-aligned and ends in a RIGHT-TO-LEFT MARK (U+200F, byte
;; FE in 1256), which keeps a final "." or ":" on the Farsi side.
;; A Farsi sentence that ends in one Latin word (a command or a layer
;; name). With a left-to-right base the Latin run is shown to the LEFT of
;; the Farsi run, i.e. where a Farsi reader finishes - so in the string it
;; has to come FIRST. (das-fa-end "... با دستور" "DAS") reads "... با دستور DAS".
(defun das-fa-end (fa tok) (strcat tok " " fa))

;; A line that starts with a LEFT-TO-RIGHT MARK (U+200E) is meant to read
;; left to right even if it holds Farsi, e.g. a folder path.
(defun das-rtl-p (s)
  (and (/= (ascii s) 8206)
       (vl-some '(lambda (c) (or (<= 1536 c 1791) (<= 64336 c 65279))) (vl-string->list s))))

(defun das-ui-message (title text buttons dflt / lines w f dcl id res n b)
  (setq lines (das-split-lines text))
  (if (> (length lines) 30)
    (setq lines (append (das-list-take lines 28)
                        (list "…" (das-tr (das-fa-end "بقیهٔ پیام‌ها در تاریخچهٔ خط فرمان است؛ کلید" "F2")
                                            "(the rest is in the command history - F2)")))))
  (setq w 40)
  (foreach l lines (setq w (max w (+ 4 (strlen l)))))
  (setq w (min w 110))
  (setq dcl (vl-filename-mktemp "das_msg" nil ".dcl")
        f   (open dcl "w"))
  (write-line (strcat "das_msg : dialog { label = \"" (das-dcl-esc title) "\";") f)
  (write-line (strcat "  : column { width = " (itoa w) ";") f)
  (foreach l lines
    (write-line (if (das-rtl-p l)
                  (strcat "    : text { label = \"" (das-dcl-esc l) (das-uchr 8207) "\"; alignment = right; }")
                  (strcat "    : text { label = \"" (das-dcl-esc l) "\"; }"))
                f))
  (write-line "  }" f)
  (write-line "  spacer;" f)
  (write-line "  : row { alignment = centered; fixed_width = true;" f)
  (setq n 0)
  (foreach b buttons
    (setq n (1+ n))
    (write-line (strcat "    : button { key = \"b" (itoa n) "\"; label = \"" (das-dcl-esc (cadr b))
                        "\"; width = 14;"
                        (if (= (car b) dflt) " is_default = true;" "")
                        (if (= n (length buttons)) " is_cancel = true;" "")
                        " }") f))
  (write-line "  }" f)
  (write-line "}" f)
  (close f)
  (setq id (load_dialog dcl) res nil)
  (if (and (> id 0) (new_dialog "das_msg" id))
    (progn
      (setq n 0)
      (foreach b buttons
        (setq n (1+ n))
        (action_tile (strcat "b" (itoa n)) (strcat "(done_dialog " (itoa (+ 10 n)) ")")))
      (setq n (start_dialog))
      (setq res (if (> n 10) (car (nth (- n 11) buttons)) dflt))
      (unload_dialog id))
    (progn
      (if (> id 0) (unload_dialog id))
      (alert (strcat title "\n\n" text))
      (setq res dflt)))
  (vl-file-delete dcl)
  res)

(defun das-ui-ok (title text)
  (das-ui-message title text (list (list "ok" (das-tr "تأیید" "OK"))) "ok"))

(defun das-ui-yes-no (title text dflt)
  (das-ui-message title text
    (list (list "yes" (das-tr "بله" "Yes")) (list "no" (das-tr "خیر" "No")))
    dflt))

;; Warnings collected while a command runs, shown together in one window
;; at the end (das-notes-show) instead of one window each. Every note is
;; also printed, so the command history keeps it.
(setq *das-notes* '())

(defun das-note (s)
  (princ s)
  (setq *das-notes* (cons (vl-string-left-trim "\n" s) *das-notes*)))

(defun das-notes-show (title / txt)
  (if *das-notes*
    (progn
      (setq txt "")
      (foreach s (reverse *das-notes*) (setq txt (strcat txt "\n" s)))
      (setq *das-notes* '())
      (das-ui-ok title txt))))

;;; ============================================================
;;; MODULE 7: DRAWING ENGINE
;;; ============================================================

(defun das-draw-markers (vertices startNum style size scale polySign smallTh smallFac styleName /
                         i p num x y pr x1 y1 x2 y2 x3 y3 x4 y4 t1 t2 t3
                         dNext dPrev localScale n distText nexts prevs
                         pNext pPrev angNext angPrev vx vy bisectAng angDiff outwardDir txtPt)
  (setq i 0 num startNum n (length vertices)
        nexts (das-list-next vertices) prevs (das-list-prev vertices))
  (if (= (type num) 'REAL) (setq num (fix num)))
  (foreach p vertices
    (setq x (car p) y (cadr p))
    (cond
      ((= style "Circle")
       (setq pr (/ size 2.0))
       (entmake (list (cons 0 "LWPOLYLINE") (cons 100 "AcDbEntity") (cons 8 "L-POINTS")
                      (cons 62 256) (cons 100 "AcDbPolyline") (cons 90 2) (cons 70 1)
                      (cons 43 size) (cons 38 0.0)
                      (cons 10 (list (- x pr) y)) (cons 42 1.0)
                      (cons 10 (list (+ x pr) y)) (cons 42 1.0))))
      ((= style "Square")
       (setq x1 (- x size) y1 (- y size) x2 (+ x size) y2 (- y size)
             x3 (+ x size) y3 (+ y size) x4 (- x size) y4 (+ y size))
       (entmake (list (cons 0 "SOLID") (cons 8 "L-POINTS") (cons 62 256)
                      (cons 10 (list x1 y1 0.0)) (cons 11 (list x2 y2 0.0))
                      (cons 12 (list x4 y4 0.0)) (cons 13 (list x3 y3 0.0)))))
      ((= style "Triangle")
       (setq t1 (list x (+ y size) 0.0)
             t2 (list (- x (* size 0.866)) (- y (* size 0.5)) 0.0)
             t3 (list (+ x (* size 0.866)) (- y (* size 0.5)) 0.0))
       (entmake (list (cons 0 "SOLID") (cons 8 "L-POINTS") (cons 62 256)
                      (cons 10 t1) (cons 11 t2) (cons 12 t3) (cons 13 t3))))
      ((= style "Xmark")
       (entmake (list (cons 0 "LINE") (cons 8 "L-POINTS") (cons 62 256)
                      (cons 10 (list (- x size) (- y size) 0.0))
                      (cons 11 (list (+ x size) (+ y size) 0.0))))
       (entmake (list (cons 0 "LINE") (cons 8 "L-POINTS") (cons 62 256)
                      (cons 10 (list (- x size) (+ y size) 0.0))
                      (cons 11 (list (+ x size) (- y size) 0.0))))))

    (setq pNext (car nexts) pPrev (car prevs)
          nexts (cdr nexts) prevs (cdr prevs))
    (setq localScale (* scale 0.8))
    (if (> smallTh 0.0)
      (progn
        (setq dNext (distance p pNext))
        (setq dPrev (distance p pPrev))
        (if (or (< dNext smallTh) (< dPrev smallTh))
          (setq localScale (* localScale smallFac)))))

    (setq angNext (angle p pNext) angPrev (angle p pPrev))
    (setq vx (+ (cos angNext) (cos angPrev)) vy (+ (sin angNext) (sin angPrev)))
    (if (and (< (abs vx) 1e-4) (< (abs vy) 1e-4))
      (setq bisectAng (+ angNext (/ pi 2.0)))
      (setq bisectAng (atan vy vx)))
    (setq angDiff (rem (+ (- angNext angPrev) (* 2.0 pi)) (* 2.0 pi)))
    ;; the interior bisector points inward when the turn direction agrees
    ;; with the polygon orientation; flip it to place the label outside
    (if (not (equal (> polySign 0) (> angDiff pi)))
      (setq outwardDir bisectAng)
      (setq outwardDir (+ bisectAng pi)))
    (setq distText (+ size (* localScale 1.5))
          txtPt    (polar p outwardDir distText))
    (das-utils-text-center txtPt localScale (itoa num) "L-TEXT" styleName)
    (setq i (1+ i) num (1+ num)))
  (princ))

;; [FIX-21] Lengths used to be aligned dimensions in whatever DIMSTYLE was
;; current. In a drawing started from acad.dwt / acadiso.dwt that is
;; Standard or ISO-25: arrowheads, a dimension line and extension lines
;; drawn over the parcel edge, text sitting above the line. A cadastral
;; sheet wants only the number, parallel to the edge, just outside it.
;; das-dim-style-ensure builds (or rebuilds) a DIMSTYLE "DAS_LEN" that
;; suppresses arrows, dimension and extension lines and centres the text
;; on the dimension line, which das-draw-dimensions puts outside the
;; parcel. The dimension variables are set only while the style is
;; copied from the document, then restored, so the user's current style
;; and its overrides are left as they were.
(setq *das-dim-style* "DAS_LEN")

(defun das-dim-style-ensure (txtH prec / doc styles st vars old ok)
  (setq doc    (vla-get-ActiveDocument (vlax-get-acad-object))
        styles (vla-get-DimStyles doc)
        vars   (list
                 (cons "DIMSCALE" 1.0)  (cons "DIMLFAC" 1.0)  (cons "DIMRND" 0.0)
                 (cons "DIMASZ"   0.0)  (cons "DIMTSZ"  0.0)
                 (cons "DIMSD1"   1)    (cons "DIMSD2"  1)
                 (cons "DIMSE1"   1)    (cons "DIMSE2"  1)
                 (cons "DIMEXO"   0.0)  (cons "DIMEXE"  0.0)
                 (cons "DIMTAD"   0)    (cons "DIMJUST" 0)    (cons "DIMTVP" 0.0)
                 (cons "DIMTIH"   0)    (cons "DIMTOH"  0)
                 (cons "DIMTIX"   1)    (cons "DIMSOXD" 1)
                 (cons "DIMATFIT" 3)    (cons "DIMTMOVE" 2)   (cons "DIMUPT" 0)
                 (cons "DIMGAP"   (* txtH 0.1))
                 (cons "DIMTXT"   txtH)
                 (cons "DIMLUNIT" 2)    (cons "DIMDEC"  prec) (cons "DIMZIN" 0)
                 (cons "DIMDSEP"  ".")  (cons "DIMPOST" "")
                 (cons "DIMCLRD"  256)  (cons "DIMCLRE" 256)  (cons "DIMCLRT" 256)))
  (if (tblsearch "STYLE" "DAS_ROMANS")
    (setq vars (cons (cons "DIMTXSTY" "DAS_ROMANS") vars)))
  ;; remember, set, copy into the style, restore
  (setq old (mapcar '(lambda (v) (cons (car v) (getvar (car v)))) vars))
  (foreach v vars (vl-catch-all-apply 'setvar (list (car v) (cdr v))))
  (setq st (vl-catch-all-apply 'vla-Item (list styles *das-dim-style*)))
  (if (vl-catch-all-error-p st)
    (setq st (vl-catch-all-apply 'vla-Add (list styles *das-dim-style*))))
  (setq ok (and st (not (vl-catch-all-error-p st))
                (not (vl-catch-all-error-p (vl-catch-all-apply 'vla-CopyFrom (list st doc))))))
  (foreach v old (vl-catch-all-apply 'setvar (list (car v) (cdr v))))
  (if ok *das-dim-style* nil))

(defun das-draw-dimensions (vertices obj offset scale polySign prec smallTh smallFac /
                            n pa pb pt1 pt2 mid ang segLen dimPt acadObj doc space dimObj
                            dimStyle txtH)
  (setvar "DIMLUNIT" 2)
  (setvar "DIMDEC" prec)
  (setq dimStyle (das-dim-style-ensure scale prec))
  (setq acadObj (vlax-get-acad-object)
        doc     (vla-get-ActiveDocument acadObj)
        space   (if (= (getvar 'cvport) 1) (vla-get-PaperSpace doc) (vla-get-ModelSpace doc))
        n       (length vertices)
        pa      vertices
        pb      (das-list-next vertices))
  (repeat (if (vlax-curve-isClosed obj) n (1- n))
    (setq pt1 (car pa)
          pt2 (car pb)
          pa  (cdr pa) pb (cdr pb)
          segLen (distance pt1 pt2)
          mid (mapcar '(lambda (a b) (/ (+ a b) 2.0)) pt1 pt2)
          ang (angle pt1 pt2)
          txtH (if (and (> smallTh 0.0) (< segLen smallTh)) (* scale smallFac) scale)
          ;; the dialog's offset is the clear gap between the edge and the
          ;; near side of the number; the text is centred on the hidden
          ;; dimension line, so that line sits half a text height further out
          dimPt (polar mid (- ang (* polySign (/ pi 2.0))) (+ offset (/ txtH 2.0))))
    (setq dimObj (vla-AddDimAligned space (vlax-3d-point pt1) (vlax-3d-point pt2) (vlax-3d-point dimPt)))
    (if dimStyle
      (vl-catch-all-apply 'vla-put-StyleName (list dimObj dimStyle))
      ;; style could not be built: switch the parts off on the object itself
      (foreach pr '((DimLine1Suppress . :vlax-true) (DimLine2Suppress . :vlax-true)
                    (ExtLine1Suppress . :vlax-true) (ExtLine2Suppress . :vlax-true))
        (if (vlax-property-available-p dimObj (car pr) T)
          (vl-catch-all-apply 'vlax-put-property (list dimObj (car pr) (eval (cdr pr)))))))
    (vla-put-Layer dimObj "L-DIM")
    (vla-put-TextHeight dimObj txtH)
    (if (vlax-property-available-p dimObj 'PrimaryUnitsPrecision)
      (vla-put-PrimaryUnitsPrecision dimObj prec)))
  (princ))

;; [FIX-2] rowsInThisCol was missing from the local list
(defun das-draw-table (ul pts obj startnum textH coordPrec withLen styleName colCount /
                       cw1 cw2 cw3 cw4 rh rhH totalW tableGap maxRows totalPts chunkStart
                       chunkEnd i idx curPt nextPt lenVal distSeg xL y0 x1 x2 x3 x4
                       cxP cxE cxN cxL xx tableEnts colOffset y rowsInThisCol rest closed)
  (setq colCount (max 1 colCount))           ; [FIX-3] guard the divisor
  (setq tableEnts '()
        cw1 (* 6.0 textH)  cw2 (* 14.0 textH) cw3 (* 14.0 textH)
        cw4 (if withLen (* 20.0 textH) 0.0)
        rh (* 2.2 textH)   rhH (* 2.6 textH)
        totalW (+ cw1 cw2 cw3 cw4)
        tableGap (* 5.0 textH)
        totalPts (length pts)
        maxRows (max 1 (/ (+ totalPts (1- colCount)) colCount))
        chunkStart 0 colOffset 0.0
        rest pts                                  ; [PERF-2] walks with idx
        closed (vlax-curve-isclosed obj))
  (while (< chunkStart totalPts)
    (setq chunkEnd (min (+ chunkStart maxRows) totalPts)
          rowsInThisCol (- chunkEnd chunkStart)
          xL (+ (car ul) colOffset) y0 (cadr ul)
          x1 (+ xL cw1) x2 (+ x1 cw2) x3 (+ x2 cw3) x4 (+ x3 cw4)
          cxP (+ xL (/ cw1 2.0)) cxE (+ x1 (/ cw2 2.0))
          cxN (+ x2 (/ cw3 2.0)) cxL (+ x3 (/ cw4 2.0)))
    (foreach xx (list xL x1 x2 x3)
      (das-utils-mkline (list xx y0 0) (list xx (- y0 rhH (* rowsInThisCol rh)) 0) "L-TABLE"))
    (if withLen
      (das-utils-mkline (list x4 y0 0) (list x4 (- y0 rhH (* rowsInThisCol rh)) 0) "L-TABLE"))
    (das-utils-mkline (list xL y0 0) (list (if withLen x4 x3) y0 0) "L-TABLE")
    (das-utils-mkline (list xL (- y0 rhH) 0) (list (if withLen x4 x3) (- y0 rhH) 0) "L-TABLE")
    (setq tableEnts (cons (das-utils-text-center (list cxP (- y0 (/ rhH 2.0)) 0) textH "POINT" "L-TABLE" styleName) tableEnts))
    (setq tableEnts (cons (das-utils-text-center (list cxE (- y0 (/ rhH 2.0)) 0) textH "EASTING" "L-TABLE" styleName) tableEnts))
    (setq tableEnts (cons (das-utils-text-center (list cxN (- y0 (/ rhH 2.0)) 0) textH "NORTHING" "L-TABLE" styleName) tableEnts))
    (if withLen
      (setq tableEnts (cons (das-utils-text-center (list cxL (- y0 (/ rhH 2.0)) 0) textH "LENGTH" "L-TABLE" styleName) tableEnts)))
    (setq i 0 idx chunkStart)
    (while (< idx chunkEnd)
      (setq y (- y0 rhH (* i rh) (/ rh 2.0)) curPt (car rest))
      (das-utils-mkline (list xL (- y0 rhH (* (1+ i) rh)) 0)
                        (list (if withLen x4 x3) (- y0 rhH (* (1+ i) rh)) 0) "L-TABLE")
      (setq tableEnts (cons (das-utils-text-center (list cxP y) textH (strcat "P" (itoa (+ startnum idx))) "L-TABLE" styleName) tableEnts))
      (setq tableEnts (cons (das-utils-text-center (list cxE y) textH (rtos (car curPt) 2 coordPrec) "L-TABLE" styleName) tableEnts))
      (setq tableEnts (cons (das-utils-text-center (list cxN y) textH (rtos (cadr curPt) 2 coordPrec) "L-TABLE" styleName) tableEnts))
      (if withLen
        (progn
          (setq nextPt (if (< idx (1- totalPts)) (cadr rest) (car pts)))
          (if (or (< idx (1- totalPts)) closed)
            (progn
              (setq distSeg (distance curPt nextPt))
              (setq lenVal (strcat "(P" (itoa (+ startnum idx)) "-P"
                                   (itoa (if (= idx (1- totalPts)) startnum (+ startnum idx 1)))
                                   ") " (rtos distSeg 2 2))))
            (setq lenVal "-"))
          (setq tableEnts (cons (das-utils-text-center (list cxL y) textH lenVal "L-TABLE" styleName) tableEnts))))
      (setq i (1+ i) idx (1+ idx) rest (cdr rest)))
    (setq chunkStart chunkEnd colOffset (+ colOffset totalW tableGap)))
  tableEnts)

;; [FIX-14] pStart/pEnd/pMid/ang below are the CHORD between a range's
;; first and last point. That is fine for a straight run, but when the
;; range bends or arcs (several in-between points, not a straight line -
;; exactly what a multi-point neighbour usually is) the real boundary can
;; bulge past that chord on the side the caption gets offset to, and a
;; fixed nOffset was not always enough to clear it - the text ended up
;; sitting on top of the polyline. perpAng/bulge below measure how far
;; the farthest in-range vertex already sits past the chord on the
;; caption's side, and add that on top of the normal offset so the label
;; clears the actual boundary, not just the chord between its endpoints.
(defun das-draw-range-neighbors (vertices neighbors nScale nOffset polySign styleName /
                                 nb sIdx eIdx txt pStart pEnd pMid ang rotAng txtPt n
                                 perpAng ux uy v bulge d extraOffset run)
  (setq n (length vertices))
  (foreach nb neighbors
    (setq sIdx (car nb) eIdx (cadr nb) txt (nth 2 nb))
    ;; [FIX-3] silently skip out-of-range indices instead of crashing
    (if (and (>= sIdx 0) (< sIdx n) (>= eIdx 0) (< eIdx n))
      (progn
        (setq pStart (nth sIdx vertices) pEnd (nth eIdx vertices))
        (setq pMid (mapcar '(lambda (a b) (/ (+ a b) 2.0)) pStart pEnd))
        (setq ang (angle pStart pEnd) rotAng ang)
        (if (and (> rotAng (/ pi 2.0)) (<= rotAng (* pi 1.5)))
          (setq rotAng (+ rotAng pi)))

        (setq perpAng (- ang (* polySign (/ pi 2.0))))
        (setq ux (cos perpAng) uy (sin perpAng))
        ;; vertices sIdx .. eIdx-1, wrapping past the last one
        (setq bulge 0.0
              run   (das-list-take (das-geom-rotate-list vertices sIdx)
                                   (rem (+ (- eIdx sIdx) n) n)))
        (foreach v run
          (setq d (+ (* (- (car v) (car pStart)) ux) (* (- (cadr v) (cadr pStart)) uy)))
          (if (> d bulge) (setq bulge d)))
        (setq extraOffset (+ nOffset bulge))

        (setq txtPt (polar pMid perpAng extraOffset))
        (entmakex
          (list (cons 0 "MTEXT") (cons 100 "AcDbEntity") (cons 8 "L-NEIGHBOR")
                (cons 100 "AcDbMText") (cons 10 txtPt) (cons 40 nScale)
                (cons 1 txt) (cons 7 styleName) (cons 50 rotAng)
                (cons 71 5) (cons 62 256))))))
  (princ))

;;; ============================================================
;;; MODULE 7B: SINGLE RENDER PATH
;;; ============================================================
;; [NEW-3] 9.6 carried two near-identical 60-line copies of this block,
;; one for the preview branch and one for the final draw. Any change had
;; to be made twice.
(defun das-render-all (vertices obj polySign withLen /
                       minx maxx miny maxy autoTablePt autoAreaPt maxLen maxAng
                       curLen p1 p2 i n tableEnts neighborsList nb txt
                       markerEnt drawnEnts bx gap areaEnt)
  (setq n (length vertices))
  (setq minx (car (car vertices))  maxx (car (car vertices))
        miny (cadr (car vertices)) maxy (cadr (car vertices)))
  (foreach p1 vertices
    (setq minx (min minx (car p1)) maxx (max maxx (car p1))
          miny (min miny (cadr p1)) maxy (max maxy (cadr p1))))

  ;; [FIX-15] area centroid first (not the bounding-box centre, which
  ;; misses far more often on a concave/notched parcel); if even the true
  ;; centroid falls outside, fall back to a point proven to be inside.
  (setq autoAreaPt (das-geom-centroid vertices))
  (if (not (das-pt-in-poly-p autoAreaPt vertices))
    (setq autoAreaPt (das-safe-label-point autoAreaPt vertices)))
  (setq autoAreaPt (list (car autoAreaPt) (cadr autoAreaPt) 0.0))

  ;; the area label follows the longest edge so it reads along the parcel
  (setq maxLen 0.0 maxAng 0.0)
  (mapcar '(lambda (p1 p2)
             (setq curLen (distance p1 p2))
             (if (> curLen maxLen) (setq maxLen curLen maxAng (angle p1 p2))))
          vertices (das-list-next vertices))
  (if (and (> maxAng (/ pi 2.0)) (<= maxAng (* pi 1.5)))
    (setq maxAng (+ maxAng pi)))
  ;; [FIX-16] steep/near-vertical looks worse than horizontal for a
  ;; summary label that has no single edge it must align with
  (if (das-angle-is-steep-p maxAng) (setq maxAng 0.0))

  ;; [FIX-12] remember where the drawing ends, so the extents of whatever
  ;; we add below can be measured before the table is positioned
  (setq markerEnt (entlast))

  (setvar 'clayer "L-POINTS")
  (das-draw-markers vertices *das-last-start* *das-last-mkstyle* *das-last-mksize*
                    *das-last-scale* polySign *das-last-small-th* *das-last-small-fac* "DAS_ROMANS")

  (setvar 'clayer "L-DIM")
  (das-draw-dimensions vertices obj *das-last-offset* *das-last-scale* polySign
                       *das-last-prec-l* *das-last-small-th* *das-last-small-fac*)

  ;; [FIX-11] "m2" -> "m" with a superscript 2
  (setvar 'clayer "L-AREA")
  (setq areaEnt
    (cond ((das-utils-area-field-mtext autoAreaPt (* 2.0 *das-last-scale*) obj
             *das-last-prec-a* "L-AREA" "DAS_ROMANS" maxAng))
          ((das-utils-mtext-rotated autoAreaPt (* 2.0 *das-last-scale*)
             (strcat "S=" (rtos (vla-get-Area obj) 2 *das-last-prec-a*) " m\\S2^;")
             "L-AREA" "DAS_ROMANS" maxAng))))
  ;; [FIX-23] size and place it so it stays inside the parcel
  (if areaEnt
    (das-area-label-place areaEnt vertices autoAreaPt maxAng
                          (* 2.0 *das-last-scale*) (* 0.8 *das-last-scale*)))

  ;; neighbours are drawn BEFORE the table now - they sit outside the
  ;; boundary and are exactly what the table used to collide with
  (if *das-temp-nb-data*
    (progn
      (setq neighborsList '())
      (foreach nb *das-temp-nb-data*
        (setq neighborsList
          (cons (list (- (car nb) *das-last-start*) (- (cadr nb) *das-last-start*) (nth 2 nb))
                neighborsList)))
      (setq neighborsList (reverse neighborsList))
      (setvar 'clayer "L-NEIGHBOR")
      (das-draw-range-neighbors vertices neighborsList *das-last-nb-scale*
                                *das-last-nb-offset* polySign "DAS_FARSI")))

  ;; now measure what is really on the sheet and drop the table under it
  (setq gap (* 5.0 *das-last-scale*))
  (setq drawnEnts (das-ents-after markerEnt))
  (setq bx (das-bbox-of drawnEnts))
  (setq autoTablePt
    (if bx
      (list (car bx) (- (cadr bx) gap) 0.0)
      (list minx (- miny gap) 0.0)))

  (setvar 'clayer "L-TABLE")
  (setq tableEnts (das-draw-table autoTablePt vertices obj *das-last-start* *das-last-scale*
                                  *das-last-prec-c* withLen "DAS_ROMANS" *das-last-col-count*))
  (if (= *das-last-mask* "Yes")
    (das-utils-apply-mask tableEnts (* 0.6 *das-last-scale*)))

  (princ))

;; [NEW-16] Repairs DAS can make on its own before QC, after asking
;; once: vertices closer than 1 cm to the previous one (duplicates from
;; snapping twice or joined lines), a last vertex lying on the first,
;; an open boundary, and a non-zero elevation. LWPOLYLINE only; the
;; polyline itself is changed (one undo step with the rest of DAS).
(setq *das-fix-tol* 0.01)

;; ed -> (header vertex-groups tail). A vertex group is the 10 code with
;; the 40/41/42/91 codes that follow it.
(defun das-fix-split (ed / hdr groups cur tail x st)
  (setq st 'hdr hdr '() groups '() cur nil tail '())
  (foreach x ed
    (cond
      ((= (car x) 10)
       (if cur (setq groups (cons (reverse cur) groups)))
       (setq cur (list x) st 'v))
      ((and (eq st 'v) (member (car x) '(40 41 42 91)))
       (setq cur (cons x cur)))
      ((eq st 'hdr) (setq hdr (cons x hdr)))
      (T (setq tail (cons x tail)))))
  (if cur (setq groups (cons (reverse cur) groups)))
  (list (reverse hdr) (reverse groups) (reverse tail)))

(defun das-fix-set42 (g b)
  (cond ((null b) g)
        ((assoc 42 g) (subst (cons 42 b) (assoc 42 g) g))
        (T (append g (list (cons 42 b))))))

(defun das-fix-pt (g) (cdr (assoc 10 g)))

;; (messages new-entity-list) or nil when nothing needs fixing.
(defun das-fix-plan (ent / ed sp hdr groups tail kept g closed nDup endDup open elev msgs new en)
  (setq ed (entget ent) en (= *das-lang* "English"))
  (if (= (cdr (assoc 0 ed)) "LWPOLYLINE")
    (progn
      (setq sp (das-fix-split ed) hdr (car sp) groups (cadr sp) tail (caddr sp)
            closed (= 1 (logand 1 (cond ((cdr (assoc 70 ed))) (0))))
            kept '() nDup 0)
      (foreach g groups
        (if (and kept (< (distance (das-fix-pt g) (das-fix-pt (car kept))) *das-fix-tol*))
          (setq nDup (1+ nDup)
                kept (cons (das-fix-set42 (car kept) (cdr (assoc 42 g))) (cdr kept)))
          (setq kept (cons g kept))))
      (setq kept (reverse kept))
      (if (and (> (length kept) 3)
               (< (distance (das-fix-pt (car kept)) (das-fix-pt (last kept))) *das-fix-tol*))
        (setq endDup T kept (reverse (cdr (reverse kept)))))
      (setq open (not closed)
            elev (and (equal (cond ((cdr (assoc 210 ed))) ('(0.0 0.0 1.0))) '(0.0 0.0 1.0) 1e-9)
                      (> (abs (cond ((cdr (assoc 38 ed))) (0.0))) 1e-6)))
      (if (and (>= (length kept) 3) (or (> nDup 0) endDup open elev))
        (progn
          (setq msgs '())
          (if (> nDup 0)
            (setq msgs (cons (if en
                               (strcat "- " (itoa nDup) " vertex(es) closer than 1 cm to the previous one: removed")
                               (strcat "- رأس تکراری یا نزدیک‌تر از یک سانتی‌متر به رأس قبلی حذف می‌شود؛ "
                                       (das-fa-num nDup) " مورد"))
                             msgs)))
          (if endDup
            (setq msgs (cons (das-tr "- رأس آخر روی رأس اول است؛ حذف می‌شود"
                                     "- the last vertex lies on the first: removed")
                             msgs)))
          (if open
            (setq msgs (cons (das-tr "- پلی‌لاین باز است؛ بسته می‌شود"
                                     "- the polyline is open: it will be closed")
                             msgs)))
          (if elev
            (setq msgs (cons (das-tr "- ارتفاع پلی‌لاین صفر نیست؛ صفر می‌شود"
                                     "- the polyline has an elevation: set to 0")
                             msgs)))
          (setq new hdr)
          (setq new (subst (cons 90 (length kept)) (assoc 90 new) new))
          (setq new (subst (cons 70 (logior 1 (cond ((cdr (assoc 70 new))) (0)))) (assoc 70 new) new))
          (if (and elev (assoc 38 new)) (setq new (subst (cons 38 0.0) (assoc 38 new) new)))
          (foreach g kept (setq new (append new g)))
          (setq new (append new tail))
          (list (reverse msgs) new))))))

;; Ask and repair. T when the polyline was changed.
(defun das-fix-run (ent / plan txt)
  (if (setq plan (das-fix-plan ent))
    (progn
      (setq txt (das-tr "این ایرادها در مرز قطعه پیدا شد:" "DAS found these problems in the boundary:"))
      (foreach m (car plan) (setq txt (strcat txt "\n" m)))
      (setq txt (strcat txt "\n\n"
                        (das-tr (das-fa-end "خودکار اصلاح شود؟ بعداً هم برمی‌گردد با دستور" "U")
                                "Repair them now? U undoes it.")))
      (if (= (das-ui-message (das-tr "DAS - اصلاح خودکار مرز" "DAS - boundary repair")
                             txt
                             (list (list "fix"  (das-tr "اصلاح" "Repair"))
                                   (list "skip" (das-tr "بدون اصلاح" "Leave as is")))
                             "fix")
             "fix")
        (if (entmod (cadr plan))
          (progn
            (entupd ent)
            (princ (das-tr "\nمرز قطعه اصلاح شد." "\nBoundary repaired."))
            T)
          (progn
            (das-note (das-tr "\nاصلاح خودکار مرز انجام نشد؛ پلی‌لاین دست نخورد."
                              "\nThe boundary could not be repaired; the polyline was left as it was."))
            nil))
        nil))
    nil))

;; [NEW-14] Shared by DAS (after the dialog) and DASUPDATE (no dialog).
;; Distinct vertices, clockwise: temp1's header requires
;; "شماره ضلع گردش در جهت ساعت". nil when fewer than 3 remain.
(defun das-base-verts (ent / v)
  (setq v (das-geom-dedupe (das-geom-get-vertices ent) 0.001))
  (cond ((< (length v) 3) nil)
        ((> (das-geom-signed-area v) 0) (reverse v))
        (T v)))

;; Apply the start rule: (vertices polySign).
(defun das-start-verts (baseVerts / vertices)
  (setq vertices (das-geom-rotate-list baseVerts
                   (das-geom-start-index baseVerts *das-start-rule*)))
  (list vertices (if (> (das-geom-signed-area vertices) 0) 1.0 -1.0)))

;; The final output of one parcel from the current globals: erase its
;; previous output, draw, tag the new entities, stamp the fingerprint and
;; save the record. Returns how many old entities were erased.
(defun das-output-final (ent obj vertices polySign / owner nOld lastEnt)
  (setq owner (das-handle ent)
        nOld  (das-out-delete owner))                     ; [NEW-13]
  (setq lastEnt (entlast))
  (das-render-all vertices obj polySign *das-last-withlen*)
  (das-out-tag (das-ents-after lastEnt) owner)
  (das-fp-stamp ent)                                      ; [NEW-9]
  (das-proj-save ent)                                     ; [FIX-18]
  nOld)

;;; ============================================================
;;; MODULE 8: MAIN CONTROLLER
;;; ============================================================
;; [FIX-2] *error* and every working variable are now declared local.
;; In 9.6 the (defun *error* ...) inside das-core replaced the global
;; error handler for the rest of the AutoCAD session.
(defun das-core (withLen / *error* das-restore ent obj baseVerts vertices markerSize
                           polySign lan-choice dialogLoop status
                           lastEntBeforeDraw previewEnts nextEnt l issues qcOk qcMarks
                           oldOsm oldClay oldCmd oldDimTxt oldDimLun oldDimDec oldDimZin
                           nOld)

  (setq oldOsm    (getvar 'osmode)
        oldClay   (getvar 'clayer)
        oldCmd    (getvar 'cmdecho)
        oldDimTxt (getvar "DIMTXT")
        oldDimLun (getvar "DIMLUNIT")
        oldDimDec (getvar "DIMDEC")
        oldDimZin (getvar "DIMZIN"))

  (defun das-restore ()
    (setvar "DIMTXT" oldDimTxt) (setvar "DIMLUNIT" oldDimLun)
    (setvar "DIMDEC" oldDimDec) (setvar "DIMZIN" oldDimZin)
    (setvar 'clayer oldClay) (setvar 'osmode oldOsm) (setvar 'cmdecho oldCmd)
    (vl-catch-all-apply 'vla-EndUndoMark
      (list (vla-get-ActiveDocument (vlax-get-acad-object))))
    (princ))

  (defun *error* (msg)
    (if previewEnts (das-clean-preview previewEnts))
    (das-restore)
    (if (not (wcmatch (strcase msg) "*BREAK*,*CANCEL*,*EXIT*,*QUIT*"))
      (princ (strcat "\nError: " msg)))
    (princ))

  (setvar 'cmdecho 0) (setvar 'osmode 0) (setvar "DIMZIN" 0)
  (vla-StartUndoMark (vla-get-ActiveDocument (vlax-get-acad-object)))

  (das-init-globals)
  (setq *das-last-withlen* withLen *das-notes* '())
  (das-utils-ensure-style "DAS_ROMANS" "romans.shx")
  (das-utils-ensure-style "DAS_FARSI"  "arial.ttf")
  (foreach l *das-layers* (das-utils-ensure-layer (car l) (cadr l)))

  (setq previewEnts '())
  (setq ent (car (entsel "\nSelect Polyline: ")))

  (cond
    ((null ent) (princ "\nNothing selected."))

    ((not (wcmatch (cdr (assoc 0 (entget ent))) "LWPOLYLINE,POLYLINE"))
     (alert (das-tr "آبجکت انتخاب‌شده پلی‌لاین نیست؛ مرز قطعه را انتخاب کنید."
                    "The selected object is not a polyline; select the parcel boundary.")))

    (T
     (das-fix-run ent)                                    ; [NEW-16]
     (setq obj       (vlax-ename->vla-object ent)
           baseVerts (das-base-verts ent))                ; distinct, clockwise

     (cond
       ((null baseVerts)
        (alert (das-tr "پلی‌لاین باید دست‌کم سه رأس مجزا داشته باشد."
                      "The polyline needs at least three distinct vertices.")))

       (T
        (setq *das-nvert* (length baseVerts))
        (setq markerSize (das-geom-calc-marker-size baseVerts))

        ;; [NEW-17] no language question on the command line any more:
        ;; the dialog opens in *das-lang* (Farsi in a new session) and has
        ;; an English / فارسی button that switches it in place

        ;; [FIX-18] applicant, neighbours, start number from this parcel;
        ;; [NEW-14] and its drawing settings - except the Length column,
        ;; which follows the command typed (DAS or DASL)
        (if (das-proj-load ent)
          (princ (das-tr "\nاطلاعات ذخیره‌شدهٔ این قطعه بارگذاری شد: متقاضی، مجاورین و تنظیمات ترسیم."
                         "\nSaved data for this parcel (applicant, neighbours, settings) was loaded.")))
        ;; [FIX-22] sizes from this parcel, not from the last one drawn
        (if (das-sizes-need-auto-p ent)
          (princ (strcat (das-tr "\nاندازه‌ها از روی ابعاد قطعه تنظیم شد؛ ارتفاع متن: "
                                 "\nSizes set from the parcel's extent; text height: ")
                         (rtos (das-auto-sizes baseVerts) 2 2))))
        (setq *das-last-withlen* withLen *das-nid-warned* nil)

        ;; [NEW-7] geometry QC before the dialog
        (setq qcOk T)
        (if (setq issues (das-qc-run ent))
          (progn
            (das-qc-clear-marks)
            (setq qcMarks (das-qc-mark issues (* 1.5 markerSize)))
            (princ (das-qc-summary issues))
            ;; [NEW-15] in a window, not on the command line
            (if (das-qc-has-error-p issues)
              (setq qcOk (= (das-ui-message
                              (das-tr "DAS - خطای هندسی" "DAS - geometry errors")
                              (strcat (das-qc-summary issues) "\n\n"
                                      (das-tr (strcat (das-fa-end "جای ایرادها با دایرهٔ قرمز مشخص شده است در لایهٔ" "L-QC")
                                                      "\nبا وجود این خطاها ادامه می‌دهید؟")
                                              "The problems are circled in red on layer L-QC.\nContinue anyway?"))
                              (list (list "yes" (das-tr "ادامه" "Continue"))
                                    (list "no"  (das-tr "توقف" "Stop")))
                              "no")
                            "yes"))
              (das-ui-ok (das-tr "DAS - هشدار هندسی" "DAS - geometry warnings")
                         (strcat (das-qc-summary issues) "\n\n"
                                 (das-tr "این موارد هشدارند و مانع ادامهٔ کار نیستند."
                                         "These are warnings only; DAS continues."))))
            (if qcOk
              (foreach l qcMarks (if (and l (entget l)) (entdel l)))
              (princ (if (= *das-lang* "English")
                "\nStopped. Fix the marked points, then run DAS again."
                (strcat "\nکار متوقف شد."
                        "\n" (das-fa-end "نقطه‌های علامت‌خورده را اصلاح کنید و دوباره اجرا کنید:" "DAS")))))))

        (if qcOk (das-notes-show "DAS"))                    ; e.g. neighbours dropped
        (setq dialogLoop qcOk)
        (while dialogLoop
          (setq status (das-ui-dcl-show markerSize))

          ;; [NEW-1] the start rule is applied on every render, so it can
          ;; be changed in the dialog and previewed immediately
          (setq vertices (das-start-verts baseVerts)
                polySign (cadr vertices)
                vertices (car vertices))

          (cond
            ((= status 2)                                   ; preview
             (das-clean-preview previewEnts)
             (setq lastEntBeforeDraw (entlast) previewEnts '())
             (das-render-all vertices obj polySign *das-last-withlen*)
             (command "_.redraw")
             (setq nextEnt (if lastEntBeforeDraw (entnext lastEntBeforeDraw) (entnext)))
             (while nextEnt
               (setq previewEnts (cons nextEnt previewEnts))
               (setq nextEnt (entnext nextEnt)))
             (getstring (if (= *das-lang* "Farsi")
                          (strcat "\n" (das-fa-end "پیش‌نمایش را بررسی کنید؛ زوم و جابه‌جایی آزاد است. برای برگشت به پنجره بزنید:" "Enter"))
                          "\n--- Review the layout (Zoom/Pan are free). Press [Enter] to return...")))

            ((= status 1)                                   ; OK
             (das-clean-preview previewEnts)
             (setq previewEnts '() dialogLoop nil)
             ;; [NEW-13] replaces this parcel's previous output (U undoes it)
             (if (> (setq nOld (das-output-final ent obj vertices polySign)) 0)
               (princ (das-tr (strcat "\nخروجی قبلی این قطعه پاک شد (" (das-fa-num nOld) " آبجکت).")
                              (strcat "\nPrevious output of this parcel erased: " (itoa nOld)))))
             (if (das-proj-complete-p)
               (progn
                 (das-notes-show "DAS")                     ; warnings first
                 (if (= (das-ui-yes-no "DAS"
                          (das-tr "گواهی و نقشه در فایل اکسل صادر شود؟"
                                  "Export the certificate and the map to Excel?")
                          "no")
                        "yes")
                   (das-export-excel-module vertices obj *das-last-start*
                                            *das-temp-nb-data* (vla-get-Area obj))))
               (das-note (das-tr (strcat "\nفایل اکسل ساخته نشد، چون نام متقاضی یا کد ملی خالی است."
                                         "\n" (das-fa-end "برای صدور، این دو را وارد کنید و دوباره اجرا کنید:" "DAS"))
                                 "\nExcel export skipped: applicant or national ID is empty. Run DAS again and fill them in to export."))))

            ((= status 3)                                   ; [NEW-17] language
             (setq *das-lang* (if (and (= *das-lang* "English") (das-farsi-ok-p)) "Farsi" "English")))

            (T                                              ; Cancel
             (das-clean-preview previewEnts)
             (setq previewEnts '() dialogLoop nil))))))))

  (das-restore)
  (das-notes-show "DAS")                                    ; [NEW-15]
  (princ (strcat "\nDAS " *das-version* " done."))
  (princ))

;;; ============================================================
;;; MODULE 9: QC AND CHANGE-TRACKING COMMANDS  [NEW-7, NEW-9, v10.3]
;;; ============================================================
;; On final output DAS stamps the polyline with XDATA (app DAS_TOOLS):
;; version, date, vertex count, area and a coordinate checksum. DASCHECK
;; recomputes them and lists parcels edited after their output was made.
;; XDATA travels with the entity through MOVE/COPY/STRETCH, so a moved
;; parcel also reads as changed - its coordinates did change.
(defun das-fp-checksum (pts / s i)
  (setq s 0.0 i 1)
  (foreach p pts
    (setq s (+ s (* i (+ (* 0.7071 (car p)) (* 1.3027 (cadr p)))))
          i (1+ i)))
  s)

(defun das-fp-of (ent / pts)
  (setq pts (das-geom-dedupe (das-geom-get-vertices ent) 0.001))
  (list (length pts) (vla-get-Area (vlax-ename->vla-object ent)) (das-fp-checksum pts)))

(defun das-fp-stamp (ent / fp ed)
  (if (not (tblsearch "APPID" "DAS_TOOLS")) (regapp "DAS_TOOLS"))
  (setq fp (das-fp-of ent)
        ed (entget ent))
  (entmod
    (append ed
      (list (list -3 (list "DAS_TOOLS"
                           (cons 1000 *das-version*)
                           (cons 1000 (menucmd "M=$(edtime,$(getvar,date),YYYY-MO-DD)"))
                           (cons 1071 (car fp))
                           (cons 1040 (cadr fp))
                           (cons 1040 (caddr fp))))))))

;; nil = never stamped, T = unchanged, otherwise the stamp date string.
(defun das-fp-status (ent / xd vals fp reals)
  (setq xd (cdr (assoc -3 (entget ent '("DAS_TOOLS")))))
  (if (null xd)
    nil
    (progn
      (setq vals  (cdr (car xd))
            reals (mapcar 'cdr (vl-remove-if-not '(lambda (x) (= (car x) 1040)) vals))
            fp    (das-fp-of ent))
      (if (and (= (car fp) (cdr (assoc 1071 vals)))
               (equal (cadr fp)  (car reals)  0.0005)
               (equal (caddr fp) (cadr reals) 0.0001))
        T
        (cdr (nth 1 (vl-remove-if-not '(lambda (x) (= (car x) 1000)) vals)))))))

;;; ============================================================
;;; MODULE 9B: PER-PARCEL PROJECT DATA, NATIONAL ID, OUTPUT TAGS
;;;            [FIX-18, NEW-12, NEW-13, v10.5]
;;; ============================================================
(defun das-tr (fa en) (if (= *das-lang* "English") en fa))

;; Counts and dates inside Farsi text use Persian digits. (Version
;; numbers, layer and command names stay Latin.) Characters outside code
;; page 1256 reach the DCL file as \U+06F1 and DCL decodes them.
(defun das-fa-digits (s)
  (if (das-unicode-p) (vl-string-translate "0123456789" "۰۱۲۳۴۵۶۷۸۹" s) s))   ; [FIX-25]
(defun das-fa-num (n) (das-fa-digits (itoa n)))

(defun das-handle (ent) (cdr (assoc 5 (entget ent))))

;; ---- project record on the polyline ([FIX-18]) ----------------------
;; Stored with vlax-ldata in the polyline's extension dictionary as a
;; list of (key value). COPY duplicates the dictionary, so the record
;; also carries the owner's handle: a copied parcel has a new handle and
;; is treated as having no record, instead of inheriting the applicant.
(defun das-proj-get (d key) (cadr (assoc key d)))

;; [NEW-14] The drawing settings of the dialog, stored with the parcel so
;; DASUPDATE can redraw it exactly as it was first drawn. Kept as
;; (name value) pairs: a record from 10.5 has none and simply keeps the
;; current session values.
(setq *das-render-vars*
  '(*das-last-scale* *das-last-offset* *das-last-prec-c* *das-last-prec-l*
    *das-last-prec-a* *das-last-mask* *das-last-mkstyle* *das-last-mksize*
    *das-last-small-th* *das-last-small-fac* *das-last-col-count*
    *das-last-nb-scale* *das-last-nb-offset* *das-last-withlen*))

;; Everything a parcel record overwrites, so DASUPDATE can put the
;; session back the way it was when it is done.
(setq *das-parcel-vars*
  (append *das-render-vars*
          '(*das-applicant* *das-nat-id* *das-b-area* *das-zone* *das-last-start*
            *das-start-rule* *das-nvert* *das-temp-nb-data* *das-temp-nb-strings*)))

(defun das-vars-get (syms)
  (mapcar '(lambda (s) (list (vl-symbol-name s) (eval s))) syms))

(defun das-vars-set (pairs)
  (foreach p pairs (set (read (car p)) (cadr p))))

(defun das-proj-save (ent)
  (vl-catch-all-apply 'vlax-ldata-put
    (list ent "DAS_PROJECT"
          (list (list "owner"      (das-handle ent))
                (list "applicant"  *das-applicant*)
                (list "nat-id"     *das-nat-id*)
                (list "b-area"     *das-b-area*)
                (list "zone"       *das-zone*)
                (list "start"      *das-last-start*)
                (list "rule"       *das-start-rule*)
                (list "nvert"      *das-nvert*)
                (list "nb-data"    *das-temp-nb-data*)
                (list "nb-strings" *das-temp-nb-strings*)
                (list "settings"   (das-vars-get *das-render-vars*))))))

;; The record, or nil when there is none or it was copied from another
;; polyline (owner handle differs).
(defun das-proj-read (ent / d)
  (setq d (vl-catch-all-apply 'vlax-ldata-get (list ent "DAS_PROJECT")))
  (if (or (vl-catch-all-error-p d) (not (listp d))
          (/= (das-proj-get d "owner") (das-handle ent)))
    nil
    d))

;; Fills the project globals, the drawing settings and the neighbour
;; lists from the record, or empties the project fields. Neighbour ranges
;; are vertex numbers, so they are only kept while the vertex count
;; (*das-nvert*, set by the caller) is unchanged.
;; Returns nil (no record), T, or 'NBDROP (record loaded, neighbours dropped).
(defun das-proj-load (ent / d v res)
  (setq d (das-proj-read ent) res T)
  (setq *das-temp-nb-data* '() *das-temp-nb-strings* '())
  (cond
    (d
     (if (das-proj-get d "settings") (das-vars-set (das-proj-get d "settings")))
     (setq *das-applicant* (cond ((das-proj-get d "applicant")) (""))
           *das-nat-id*    (cond ((das-proj-get d "nat-id")) (""))
           *das-b-area*    (cond ((numberp (setq v (das-proj-get d "b-area"))) v) (0.0)))
     (if (numberp (setq v (das-proj-get d "zone")))  (setq *das-zone* v))
     (if (numberp (setq v (das-proj-get d "start"))) (setq *das-last-start* v))
     (if (member (setq v (das-proj-get d "rule")) '("NorthWest" "Northernmost" "AsDrawn"))
       (setq *das-start-rule* v))
     (if (= (das-proj-get d "nvert") *das-nvert*)
       (setq *das-temp-nb-data*    (das-proj-get d "nb-data")
             *das-temp-nb-strings* (das-proj-get d "nb-strings"))
       (if (das-proj-get d "nb-data")
         (progn
           (setq res 'NBDROP)
           (das-note (das-tr (strcat "\nتعداد رئوس قطعه تغییر کرده است، پس فهرست قبلی مجاورین حذف شد."
                                     "\n" (das-fa-end "مجاورین را دوباره وارد کنید با دستور" "DAS"))
                             (strcat "\nThe vertex count changed, so the saved neighbour list was dropped; enter it again with DAS. Handle: "
                                     (das-handle ent)))))))
     res)
    (T
     (setq *das-applicant* "" *das-nat-id* "" *das-b-area* 0.0)
     nil)))

(defun das-proj-complete-p ()
  (and (/= *das-applicant* "") (/= *das-nat-id* "")))

;; ---- national ID ([NEW-12]) -----------------------------------------
;; Persian (U+06F0..) and Arabic (U+0660..) digits -> ASCII; blanks and
;; dashes dropped, so "۰۰۱-۲۳۴۵۶۷-۸" is accepted as typed.
(defun das-digits-ascii (s / out i c a)
  (setq out "" i 0)
  (repeat (strlen s)
    (setq i (1+ i) c (substr s i 1) a (ascii c))
    (setq out (strcat out
                (cond ((<= 1776 a 1785) (chr (- a 1728)))
                      ((<= 1632 a 1641) (chr (- a 1584)))
                      ((member c '(" " "-")) "")
                      (T c)))))
  out)

;; 10 digits: personal code. 8-9 digits are codes whose leading zeros
;; were dropped. Check digit c = last digit, s = sum d[i]*(10-i) for the
;; first nine, r = s mod 11: valid when c = r (r < 2) or c = 11 - r.
;; A run of one repeated digit passes the formula but is never issued.
;; 11 digits: legal-entity ID. k = d[9] + 2, s = sum (d[i] + k) * w[i]
;; with w = 29 27 23 19 17 repeated, r = s mod 11 (10 -> 0), r = d[10].
(defun das-national-id-valid-p (s / d n i sum r k w)
  (setq n (strlen s))
  (cond
    ((or (= n 0) (not (vl-every '(lambda (ch) (<= 48 ch 57)) (vl-string->list s)))) nil)
    ((<= 8 n 10)
     (while (< (strlen s) 10) (setq s (strcat "0" s)))
     (setq d (mapcar '(lambda (ch) (- ch 48)) (vl-string->list s)))
     (if (vl-every '(lambda (x) (= x (car d))) d)
       nil
       (progn
         (setq i 0 sum 0)
         (repeat 9 (setq sum (+ sum (* (nth i d) (- 10 i))) i (1+ i)))
         (setq r (rem sum 11))
         (if (< r 2) (= (nth 9 d) r) (= (nth 9 d) (- 11 r))))))
    ((= n 11)
     (setq d (mapcar '(lambda (ch) (- ch 48)) (vl-string->list s))
           k (+ (nth 9 d) 2)
           w '(29 27 23 19 17 29 27 23 19 17)
           i 0 sum 0)
     (repeat 10 (setq sum (+ sum (* (+ (nth i d) k) (nth i w))) i (1+ i)))
     (setq r (rem sum 11))
     (if (= r 10) (setq r 0))
     (= r (nth 10 d)))
    (T nil)))

;; ---- output ownership ([NEW-13]) ------------------------------------
;; Each entity of the final output gets XDATA DAS_OUT with a 1005 handle
;; to its parcel. 1005 (not a plain string) is translated by COPY when the
;; parcel and its annotation are copied together, so a copied set points
;; at the copied parcel and is not erased by a re-run on the original.
;; SetXData is used instead of entmod: entmod on the area MTEXT would
;; write its evaluated text back and break the FIELD.
(defun das-out-tag (ents owner / types vals)
  (if (not (tblsearch "APPID" "DAS_OUT")) (regapp "DAS_OUT"))
  (setq types (vlax-make-safearray vlax-vbInteger '(0 . 1))
        vals  (vlax-make-safearray vlax-vbVariant '(0 . 1)))
  (vlax-safearray-fill types '(1001 1005))
  (vlax-safearray-fill vals (list (vlax-make-variant "DAS_OUT") (vlax-make-variant owner)))
  (foreach e ents
    (vl-catch-all-apply
      (function (lambda () (vla-SetXData (vlax-ename->vla-object e) types vals)))
      '()))
  (princ))

(defun das-out-find (owner / ss i e xd out)
  (setq out '())
  (if (setq ss (ssget "_X" '((-3 ("DAS_OUT")))))
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq e  (ssname ss i) i (1+ i)
              xd (cdr (cadr (assoc -3 (entget e '("DAS_OUT"))))))
        (if (= (cdr (assoc 1005 xd)) owner) (setq out (cons e out))))))
  out)

;; Returns how many were erased (entities on a locked layer are left).
(defun das-out-delete (owner / n)
  (setq n 0)
  (foreach e (das-out-find owner)
    (if (not (vl-catch-all-error-p (vl-catch-all-apply 'entdel (list e))))
      (if (null (entget e)) (setq n (1+ n)))))
  n)

(defun c:dasqc (/ ss i ent issues nErr nWarn en)
  (das-init-globals)
  (setq en (= *das-lang* "English") *das-notes* '())
  (princ (if en "\nSelect parcels to check: " "\nقطعه‌ها را برای بررسی انتخاب کنید: "))
  (if (setq ss (ssget '((0 . "LWPOLYLINE,POLYLINE"))))
    (progn
      (das-qc-clear-marks)
      (setq i 0 nErr 0 nWarn 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i) issues (das-qc-run ent))
        (if issues
          (progn
            (das-qc-mark issues
              (* 1.5 (das-geom-calc-marker-size (das-geom-dedupe (das-geom-get-vertices ent) 0.001))))
            (das-note (strcat (if en
                                (strcat "\nPolyline, handle " (das-handle ent) ":")
                                (strcat "\nپلی‌لاین " (das-fa-num (1+ i)) " از انتخاب:"))
                              (das-qc-summary issues) "\n "))
            (if (das-qc-has-error-p issues) (setq nErr (1+ nErr)) (setq nWarn (1+ nWarn)))))
        (setq i (1+ i)))
      ;; [NEW-15] counts that are not zero, one per line, Farsi first
      (das-note (das-tr (strcat "\nتعداد قطعه‌های بررسی‌شده: " (das-fa-num (sslength ss)))
                        (strcat "\nParcels checked: " (itoa (sslength ss)))))
      (if (> nErr 0)
        (das-note (das-tr (strcat "\nقطعه‌های دارای خطا: " (das-fa-num nErr))
                          (strcat "\nWith errors: " (itoa nErr)))))
      (if (> nWarn 0)
        (das-note (das-tr (strcat "\nقطعه‌های فقط با هشدار: " (das-fa-num nWarn))
                          (strcat "\nWith warnings only: " (itoa nWarn)))))
      (das-note (if (> (+ nErr nWarn) 0)
                  (das-tr (strcat "\n" (das-fa-end "جای ایرادها با دایره مشخص شده است در لایهٔ" "L-QC"))
                          "\nThe problems are circled on layer L-QC.")
                  (das-tr "\nایرادی پیدا نشد." "\nNo problems found.")))
      (das-notes-show "DASQC")))                            ; [NEW-15]
  (princ))

(defun c:dascheck (/ ss i ent st nOk nChg nNone en c)
  (das-init-globals)
  (setq en (= *das-lang* "English") *das-notes* '())
  (princ (if en "\nSelect parcels <Enter = all DAS parcels>: "
                "\nقطعه‌ها را انتخاب کنید <Enter = همهٔ قطعه‌های DAS>: "))
  (setq ss (ssget '((0 . "LWPOLYLINE,POLYLINE"))))
  (if (null ss)
    (setq ss (ssget "_X" '((0 . "LWPOLYLINE,POLYLINE") (-3 ("DAS_TOOLS"))))))
  (if ss
    (progn
      (das-qc-clear-marks)
      (setq i 0 nOk 0 nChg 0 nNone 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i) st (das-fp-status ent))
        (cond
          ((null st) (setq nNone (1+ nNone)))
          ((= st T)  (setq nOk (1+ nOk)))
          (T
           (setq nChg (1+ nChg)
                 c (das-geom-centroid (das-geom-dedupe (das-geom-get-vertices ent) 0.001)))
           (das-qc-mark (list (list 'chg (list (car c) (cadr c) 0.0)))
             (* 3.0 (das-geom-calc-marker-size (das-geom-dedupe (das-geom-get-vertices ent) 0.001))))
           ;; date as 2026/10/16: "/" keeps the digits together in a Farsi
           ;; line, "-" would let them come out reversed
           (das-note (das-tr (strcat "\nیک قطعه بعد از خروجی تاریخ "
                                     (das-fa-digits (vl-string-translate "-" "/" st)) " تغییر کرده است.")
                             (strcat "\nChanged since the DAS output of " st ". Handle: "
                                     (das-handle ent))))))
        (setq i (1+ i)))
      ;; [NEW-15] counts that are not zero, one per line, Farsi first
      (if (> nChg 0) (das-note "\n "))
      (if (> nChg 0)
        (das-note (das-tr (strcat "\nتعداد قطعه‌های تغییرکرده: " (das-fa-num nChg)
                                  "\n" (das-fa-end "این قطعه‌ها با دایره مشخص شده‌اند در لایهٔ" "L-QC"))
                          (strcat "\nChanged (circled on layer L-QC): " (itoa nChg)))))
      (if (> nOk 0)
        (das-note (das-tr (strcat "\nقطعه‌های بدون تغییر: " (das-fa-num nOk))
                          (strcat "\nUnchanged: " (itoa nOk)))))
      (if (> nNone 0)
        (das-note (das-tr (strcat "\nقطعه‌های بدون سابقهٔ خروجی در نسخهٔ 10.3 به بعد: " (das-fa-num nNone))
                          (strcat "\nNever output by DAS 10.3+: " (itoa nNone))))))
    (das-note (das-tr "\nقطعه‌ای با سابقهٔ خروجی پیدا نشد." "\nNo DAS parcels found.")))
  (if (and ss (> nChg 0))
    (das-note (das-tr (strcat "\n" (das-fa-end "برای بازترسیم قطعه‌های تغییرکرده، این دستور را اجرا کنید:" "DASUPDATE"))
                      "\nRun DASUPDATE to redraw the changed parcels.")))
  (das-notes-show "DASCHECK")                               ; [NEW-15]
  (princ))

;; [NEW-14] Redraw edited parcels without the dialog: vertex numbers,
;; dimensions, area label, neighbours and the coordinate table come back
;; from the parcel's own record (settings, start rule, neighbours), the
;; old output is erased via its DAS_OUT tags and the fingerprint is
;; renewed. A selection is redrawn as picked; Enter takes every parcel in
;; the current space whose fingerprint no longer matches. Parcels with
;; geometry ERRORs are not redrawn (circled on L-QC instead), and parcels
;; with no record (never drawn by DAS 10.5+) are left for DAS. Each parcel
;; starts from the session's own settings, so a 10.5 record (no settings)
;; never inherits the previous parcel's. Excel files are not touched.
(defun c:dasupdate (/ *error* das-upd-restore saved ss all i ent base vv issues en st
                      nDone nErr nNoRec nSame nNb
                      oldOsm oldClay oldCmd oldDimTxt oldDimLun oldDimDec oldDimZin)
  (setq oldOsm    (getvar 'osmode)
        oldClay   (getvar 'clayer)
        oldCmd    (getvar 'cmdecho)
        oldDimTxt (getvar "DIMTXT")
        oldDimLun (getvar "DIMLUNIT")
        oldDimDec (getvar "DIMDEC")
        oldDimZin (getvar "DIMZIN"))
  (das-init-globals)
  (setq saved (das-vars-get *das-parcel-vars*)
        en    (= *das-lang* "English")
        *das-notes* '())

  (defun das-upd-restore ()
    (das-vars-set saved)
    (setvar "DIMTXT" oldDimTxt) (setvar "DIMLUNIT" oldDimLun)
    (setvar "DIMDEC" oldDimDec) (setvar "DIMZIN" oldDimZin)
    (setvar 'clayer oldClay) (setvar 'osmode oldOsm) (setvar 'cmdecho oldCmd)
    (vl-catch-all-apply 'vla-EndUndoMark
      (list (vla-get-ActiveDocument (vlax-get-acad-object))))
    (princ))

  (defun *error* (msg)
    (das-upd-restore)
    (if (not (wcmatch (strcase msg) "*BREAK*,*CANCEL*,*EXIT*,*QUIT*"))
      (princ (strcat "\nError: " msg)))
    (princ))

  (setvar 'cmdecho 0) (setvar 'osmode 0) (setvar "DIMZIN" 0)
  (vla-StartUndoMark (vla-get-ActiveDocument (vlax-get-acad-object)))
  (das-utils-ensure-style "DAS_ROMANS" "romans.shx")
  (das-utils-ensure-style "DAS_FARSI"  "arial.ttf")
  (foreach l *das-layers* (das-utils-ensure-layer (car l) (cadr l)))

  (princ (if en "\nSelect parcels to redraw <Enter = all changed parcels>: "
                "\nقطعه‌ها را برای بازترسیم انتخاب کنید <Enter = همهٔ قطعه‌های تغییرکرده>: "))
  (setq all nil ss (ssget '((0 . "LWPOLYLINE,POLYLINE"))))
  (if (null ss)
    (setq all T
          ss  (ssget "_X" (list '(0 . "LWPOLYLINE,POLYLINE") (cons 410 (getvar "CTAB"))
                                '(-3 ("DAS_TOOLS"))))))
  (setq nDone 0 nErr 0 nNoRec 0 nSame 0 nNb 0)
  (if ss
    (progn
      (das-qc-clear-marks)
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i) i (1+ i))
        (das-vars-set saved)
        ;; [NEW-15] one plain sentence per parcel saying what happened.
        ;; The Farsi lines carry no handle (a Latin token inside a Farsi
        ;; line scrambles it, see das-fa-end); the English ones do.
        (cond
          ((null (das-proj-read ent))
           (setq nNoRec (1+ nNoRec))
           (if (not all)
             (das-note (das-tr (strcat "\nاین قطعه با نسخهٔ 10.5 یا بعد از آن ترسیم نشده است و بازترسیم نشد."
                                       "\n" (das-fa-end "اول آن را ترسیم کنید با دستور" "DAS"))
                               (strcat "\nThis parcel was not drawn by DAS 10.5 or later; draw it with DAS first. Handle: "
                                       (das-handle ent))))))
          ((and all (= (setq st (das-fp-status ent)) T))
           (setq nSame (1+ nSame)))
          ((or (null (setq base (das-base-verts ent)))
               (das-qc-has-error-p (setq issues (das-qc-run ent))))
           (setq nErr (1+ nErr))
           (if (null base) (setq issues (das-qc-run ent)))
           (das-qc-mark issues
             (* 1.5 (das-geom-calc-marker-size (das-geom-dedupe (das-geom-get-vertices ent) 0.001))))
           (das-note (strcat (das-tr (strcat "\nیک قطعه بازترسیم نشد، چون شکلش خطای هندسی دارد."
                                             "\n" (das-fa-end "جای ایراد با دایره مشخص شده است در لایهٔ" "L-QC"))
                                     (strcat "\nParcel not redrawn: its shape has a geometry error, circled on layer L-QC. Handle: "
                                             (das-handle ent)))
                             (das-qc-summary issues)
                             "\n ")))
          (T
           (if (not all) (setq st (das-fp-status ent)))
           (das-note (if (= st T)
                       (das-tr "\nیک قطعه دوباره ترسیم شد؛ شکلش تغییری نکرده بود."
                               (strcat "\nParcel redrawn; its shape had not changed. Handle: " (das-handle ent)))
                       (das-tr "\nیک قطعه تغییر کرده بود و ترسیمش به‌روز شد."
                               (strcat "\nParcel had changed; its drawing was updated. Handle: " (das-handle ent)))))
           (setq *das-nvert* (length base))
           (if (eq (das-proj-load ent) 'NBDROP)             ; notes it itself
             (setq nNb (1+ nNb)))
           (setq vv (das-start-verts base))
           (das-output-final ent (vlax-ename->vla-object ent) (car vv) (cadr vv))
           (setq nDone (1+ nDone)))))))

  (das-upd-restore)
  ;; [NEW-15] summary: only the counts that are not zero, one per line
  (if (> (+ nDone nErr nNoRec) 0) (das-note "\n "))
  (if (> nDone 0)
    (das-note (das-tr (strcat "\nتعداد قطعه‌های به‌روزشده: " (das-fa-num nDone))
                      (strcat "\nParcels updated: " (itoa nDone)))))
  (if (> nErr 0)
    (das-note (das-tr (strcat "\nبازترسیم‌نشده به‌خاطر خطای هندسی: " (das-fa-num nErr))
                      (strcat "\nNot redrawn because of geometry errors: " (itoa nErr)))))
  (if (and all (> nNoRec 0))
    (das-note (das-tr (strcat "\nقطعه‌های بدون سابقهٔ ترسیم در نسخهٔ 10.5 به بعد: " (das-fa-num nNoRec))
                      (strcat "\nSkipped, no DAS 10.5+ record: " (itoa nNoRec)))))
  (if (= (+ nDone nErr nNoRec) 0)
    (das-note (if all
                (das-tr "\nقطعهٔ تغییرکرده‌ای پیدا نشد؛ ترسیم همهٔ قطعه‌ها به‌روز است."
                        "\nNo changed parcels found; every parcel matches its drawing.")
                (das-tr "\nهیچ قطعه‌ای بازترسیم نشد." "\nNo parcel was redrawn."))))
  (if (> nDone 0)
    (das-note (das-tr (strcat "\nفایل‌های اکسل تغییر نکرده‌اند."
                              "\n" (das-fa-end "اگر لازم است، دوباره صادرشان کنید با دستور" "DAS"))
                      "\nExcel files were not changed; export again with DAS if needed.")))
  (das-notes-show "DASUPDATE")
  (princ))

;;; ============================================================
;;; MODULE 10: DASIMP  -  ورود نقاط از فایل CSV / TXT   [NEW-10]
;;; ============================================================
;; Pipeline: read file -> strip BOM -> detect delimiter -> detect header
;; -> map columns (by header names, else by looking at the data) ->
;; classify each point (name prefix, then Code column, else ask) ->
;; text summary + on-screen preview (grdraw, nothing entmake'd) ->
;; confirm -> draw POINT + TEXT per category layer.
;; The pure parsing helpers (das-imp-split ... das-imp-classify) never
;; touch the drawing, so they can be tested by hand with (das-imp-load).

;; ---- small utilities -----------------------------------------
(defun das-imp-take (lst n / out)
  (setq out '())
  (while (and lst (> n 0))
    (setq out (cons (car lst) out) lst (cdr lst) n (1- n)))
  (reverse out))

(defun das-imp-put (lst idx val / i out e)
  (setq i 0 out '())
  (foreach e lst
    (setq out (cons (if (= i idx) val e) out) i (1+ i)))
  (reverse out))

(defun das-imp-join (lst sep / out e)
  (setq out "")
  (foreach e lst
    (setq out (if (= out "") (vl-princ-to-string e)
                  (strcat out sep (vl-princ-to-string e)))))
  out)

(defun das-imp-in-e-p (v) (and v (>= v 160000.0) (<= v 840000.0)))
(defun das-imp-in-n-p (v) (and v (>= v 2600000.0) (<= v 4500000.0)))

;; ---- file reading --------------------------------------------
;; A UTF-8 BOM arrives as three characters (239 187 191) or, on newer
;; AutoCAD, as one character (65279). A UTF-16 file starts with 255 254
;; or 254 255 and cannot be read line by line - the caller is told.
(defun das-imp-bom-kind (s / a b c)
  (if (and s (> (strlen s) 0))
    (progn
      (setq a (ascii (substr s 1 1))
            b (if (> (strlen s) 1) (ascii (substr s 2 1)) 0)
            c (if (> (strlen s) 2) (ascii (substr s 3 1)) 0))
      (cond
        ((= a 65279) "utf8-1")
        ((and (= a 239) (= b 187) (= c 191)) "utf8-3")
        ((or (and (= a 255) (= b 254)) (and (= a 254) (= b 255))) "utf16")
        (T nil)))
    nil))

(defun das-imp-read-lines (path / f ln out first kind)
  (setq out '() first T kind nil)
  (if (setq f (open path "r"))
    (progn
      (while (and (not (equal kind "utf16")) (setq ln (read-line f)))
        (if first
          (progn
            (setq first nil kind (das-imp-bom-kind ln))
            (cond ((equal kind "utf8-1") (setq ln (substr ln 2)))
                  ((equal kind "utf8-3") (setq ln (substr ln 4))))))
        (if (not (equal kind "utf16")) (setq out (cons ln out))))
      (close f)
      (if (equal kind "utf16") nil (reverse out)))
    nil))

;; ---- splitting and numbers -----------------------------------
;; delim is "," ";" "\t" or " ". With " " a run of blanks/tabs counts
;; as ONE separator. Double quotes protect a delimiter inside a field.
(defun das-imp-split (line delim / out cur i n ch inq ws)
  (setq out '() cur "" i 0 n (strlen line) inq nil ws (= delim " "))
  (while (< i n)
    (setq i (1+ i) ch (substr line i 1))
    (cond
      ((= ch "\"") (setq inq (not inq)))
      ((and (not inq) (or (= ch delim) (and ws (= ch "\t"))))
       (if (not (and ws (= cur "")))
         (setq out (cons (vl-string-trim " \t" cur) out)))
       (setq cur ""))
      (T (setq cur (strcat cur ch)))))
  (if (not (and ws (= cur "")))
    (setq out (cons (vl-string-trim " \t" cur) out)))
  (reverse out))

;; Strict: optional sign, digits, at most one "." or ",". Letters (even
;; "E") make it non-numeric, so a point called E1 is never read as 1.
(defun das-imp-num-p (s / n i c dig sep bad)
  (setq s (vl-string-trim " \t" s) n (strlen s) i 0 dig 0 sep 0 bad nil)
  (if (and (> n 0) (member (substr s 1 1) '("+" "-"))) (setq i 1))
  (while (and (not bad) (< i n))
    (setq c (substr s (1+ i) 1) i (1+ i))
    (cond ((wcmatch c "#") (setq dig (1+ dig)))
          ((member c '("." ",")) (setq sep (1+ sep)))
          (T (setq bad T))))
  (and (> n 0) (not bad) (> dig 0) (<= sep 1)))

;; decimal comma ("500000,12") -> decimal point
(defun das-imp-num (s)
  (atof (vl-string-translate "," "." (vl-string-trim " \t" s))))

(defun das-imp-mode-count (counts / best bestf c f)
  (setq best 0 bestf 0)
  (foreach c counts
    (setq f (length (vl-remove-if-not '(lambda (x) (= x c)) counts)))
    (if (> f bestf) (setq best c bestf f)))
  (list best bestf))

;; Tab and ";" are tried before "," because a ";" file with decimal
;; commas ("B1;500000,12;3500000,55") would otherwise look like a
;; consistent 3-column comma file. A candidate wins when >= 80 % of the
;; first 30 non-blank lines split into the same number (>= 2) of fields.
(defun das-imp-detect-delim (lines / sample cand res out ln)
  (setq sample '())
  (foreach ln lines
    (if (and (< (length sample) 30) (/= (vl-string-trim " \t" ln) ""))
      (setq sample (cons ln sample))))
  (setq sample (reverse sample) out nil)
  (foreach cand '("\t" ";" "," " ")
    (if (null out)
      (progn
        (setq res (das-imp-mode-count
                    (mapcar '(lambda (l) (length (das-imp-split l cand))) sample)))
        (if (and (>= (car res) 2) (>= (* 10 (cadr res)) (* 8 (length sample))))
          (setq out cand)))))
  out)

;; Returns (delim header-fields-or-nil rows) where each row is
;; (line-number . fields). Lines starting with # or // are comments.
(defun das-imp-load (path / lines delim ln ln-no rows t1 hdr)
  (setq lines (das-imp-read-lines path))
  (setq delim (if lines (das-imp-detect-delim lines) nil))
  (if (null delim)
    nil
    (progn
      (setq rows '() ln-no 0)
      (foreach ln lines
        (setq ln-no (1+ ln-no) t1 (vl-string-trim " \t" ln))
        (if (and (/= t1 "") (/= (substr t1 1 1) "#") (/= (substr t1 1 2) "//"))
          (setq rows (cons (cons ln-no (das-imp-split ln delim)) rows))))
      (setq rows (reverse rows) hdr nil)
      (if (and rows
               (>= (length (cdr (car rows))) 2)
               (not (vl-some 'das-imp-num-p (cdr (car rows)))))
        (setq hdr (cdr (car rows)) rows (cdr rows)))
      (list delim hdr rows))))

;; ---- column mapping ------------------------------------------
;; A map is (name x y z code) of 0-based column indices, nil = absent.
(defun das-imp-norm (s / out i c)
  (setq out "" i 0 s (strcase (vl-string-trim " \t\"" s)))
  (repeat (strlen s)
    (setq i (1+ i) c (substr s i 1))
    (if (not (member c '(" " "_" "." "-")))
      (setq out (strcat out c))))
  out)

(defun das-imp-map-header (hdr / r i n h)
  (setq r (list nil nil nil nil nil) i 0)
  (foreach h hdr
    (setq n (das-imp-norm h))
    (cond
      ((and (null (nth 1 r)) (member n '("X" "E" "EAST" "EASTING" "EASTINGS" "XEAST")))
       (setq r (das-imp-put r 1 i)))
      ((and (null (nth 2 r)) (member n '("Y" "N" "NORTH" "NORTHING" "NORTHINGS" "YNORTH")))
       (setq r (das-imp-put r 2 i)))
      ((and (null (nth 3 r)) (member n '("Z" "H" "ELEV" "ELEVATION" "HEIGHT" "ALT" "ALTITUDE" "LEVEL")))
       (setq r (das-imp-put r 3 i)))
      ((and (null (nth 4 r)) (member n '("CODE" "DESC" "DESCRIPTION" "FEATURE" "FC" "TYPE"
                                         "LAYER" "CAT" "CATEGORY" "REMARK" "REMARKS" "NOTE" "NOTES" "ATTR")))
       (setq r (das-imp-put r 4 i)))
      ((and (null (nth 0 r)) (member n '("NAME" "NAMES" "PT" "PTNAME" "POINT" "POINTID" "POINTNAME"
                                         "PTNO" "PNT" "ID" "NO" "NUM" "NUMBER" "PTID" "POINTNO"
                                         "STATION" "STN")))
       (setq r (das-imp-put r 0 i))))
    (setq i (1+ i)))
  r)

;; per column: (all-numeric  all-integer  median)
(defun das-imp-col-stats (sample ncol / c out strs isnum isint vals)
  (setq c 0 out '())
  (repeat ncol
    (setq strs  (mapcar '(lambda (r) (if (< c (length r)) (nth c r) "")) sample)
          isnum (and strs (vl-every 'das-imp-num-p strs))
          isint nil vals nil)
    (if isnum
      (progn
        (setq isint (vl-every '(lambda (s) (not (or (vl-string-search "." s)
                                                    (vl-string-search "," s))))
                              strs))
        (setq vals (vl-sort (mapcar 'das-imp-num strs) '<))))
    (setq out (cons (list isnum isint (if vals (nth (/ (length vals) 2) vals) nil)) out)
          c (1+ c)))
  (reverse out))

;; No usable header: decide from the data. Column 0 is a point name when
;; it is text, or when it is an integer that is NOT in the E/N range of
;; Iran and two numeric columns follow it. Otherwise X,Y start at 0.
(defun das-imp-map-data (rows ncol / sample st s0 name x y z ci)
  (setq sample (mapcar 'cdr (das-imp-take rows 50)))
  (if (< ncol 2)
    nil
    (progn
      (setq st (das-imp-col-stats sample ncol) s0 (nth 0 st))
      (setq name
        (cond ((not (car s0)) 0)
              ((or (das-imp-in-e-p (caddr s0)) (das-imp-in-n-p (caddr s0))) nil)
              ((and (>= ncol 3) (cadr s0) (car (nth 1 st)) (car (nth 2 st))) 0)
              (T nil)))
      (setq x (if name 1 0) y (1+ x))
      (if (and (< y ncol) (car (nth x st)) (car (nth y st)))
        (progn
          (setq z  (if (and (< (1+ y) ncol) (car (nth (1+ y) st))) (1+ y) nil)
                ci (if z (+ y 2) (1+ y)))
          (list name x y z (if (< ci ncol) ci nil)))
        nil))))

(defun das-imp-auto-map (hdr rows / m)
  (setq m (if hdr (das-imp-map-header hdr) nil))
  (if (or (null m) (null (nth 1 m)) (null (nth 2 m)))
    (setq m (das-imp-map-data rows (if rows (length (cdr (car rows))) 0))))
  m)

(defun das-imp-col-median (rows idx / vals r)
  (setq vals '())
  (foreach r rows
    (if (and (< idx (length (cdr r))) (das-imp-num-p (nth idx (cdr r))))
      (setq vals (cons (das-imp-num (nth idx (cdr r))) vals))))
  (if vals (nth (/ (length vals) 2) (vl-sort vals '<)) nil))

;; If the "X" column holds northings and the "Y" column eastings (UTM
;; Iran), swap them. Returns (map swapped-flag).
(defun das-imp-maybe-swap (rows m / sample mx my)
  (setq sample (das-imp-take rows 200)
        mx (das-imp-col-median sample (nth 1 m))
        my (das-imp-col-median sample (nth 2 m)))
  (if (and mx my (das-imp-in-n-p mx) (das-imp-in-e-p my))
    (list (das-imp-put (das-imp-put m 1 (nth 2 m)) 2 (nth 1 m)) T)
    (list m nil)))

;; ---- records and categories ----------------------------------
;; record = (name x y z code line-no category)
(defun das-imp-build (rows m / recs bad n r f nm ix iy iz ic)
  (setq recs '() bad '() n 0
        nm (nth 0 m) ix (nth 1 m) iy (nth 2 m) iz (nth 3 m) ic (nth 4 m))
  (foreach r rows
    (setq f (cdr r) n (1+ n))
    (if (and (< ix (length f)) (< iy (length f))
             (das-imp-num-p (nth ix f)) (das-imp-num-p (nth iy f)))
      (setq recs
        (cons (list
                (if (and nm (< nm (length f)) (/= (nth nm f) "")) (nth nm f) (itoa n))
                (das-imp-num (nth ix f))
                (das-imp-num (nth iy f))
                (if (and iz (< iz (length f)) (das-imp-num-p (nth iz f)))
                  (das-imp-num (nth iz f)) 0.0)
                (if (and ic (< ic (length f))) (nth ic f) "")
                (car r)
                nil)
              recs))
      (setq bad (cons (car r) bad))))
  (list (reverse recs) (reverse bad)))

;; leading ASCII letters: "b12" -> "B", "CP-3" -> "CP", "12" -> ""
(defun das-imp-prefix-of (s / out i c)
  (setq s (strcase (vl-string-trim " \t" s)) out "" i 0)
  (while (and (< i (strlen s)) (wcmatch (setq c (substr s (1+ i) 1)) "@"))
    (setq out (strcat out c) i (1+ i)))
  out)

;; point name first, then the Code column; nil when neither matches
(defun das-imp-classify (rec / cat)
  (setq cat (assoc (das-imp-prefix-of (nth 0 rec)) *das-imp-cats*))
  (if (null cat)
    (setq cat (assoc (das-imp-prefix-of (nth 4 rec)) *das-imp-cats*)))
  (if cat (car cat) nil))

(defun das-imp-assign (recs)
  (mapcar '(lambda (r) (das-imp-put r 6 (das-imp-classify r))) recs))

(defun das-imp-cat-color (cat)
  (if cat (caddr (assoc cat *das-imp-cats*)) 7))

(defun das-imp-count-cat (recs cat)
  (length (vl-remove-if-not '(lambda (r) (equal (nth 6 r) cat)) recs)))

(defun das-imp-count-dups (keys / sorted n prev k)
  (setq sorted (vl-sort keys '<) n 0 prev nil)
  (foreach k sorted
    (if (and prev (= k prev)) (setq n (1+ n)))
    (setq prev k))
  n)

(defun das-imp-extent (recs / minx miny maxx maxy r)
  (setq minx 1.0e99 miny 1.0e99 maxx -1.0e99 maxy -1.0e99)
  (foreach r recs
    (setq minx (min minx (nth 1 r)) maxx (max maxx (nth 1 r))
          miny (min miny (nth 2 r)) maxy (max maxy (nth 2 r))))
  (list minx miny maxx maxy))

;; ---- interaction ---------------------------------------------
(defun das-imp-colstr (idx) (if idx (itoa (1+ idx)) "-"))

(defun das-imp-report (path delim hdr m swapped recs bad / cfg n unc outR)
  (princ "\n--- DASIMP ---")
  (princ (strcat "\nفایل: " (vl-filename-base path) (or (vl-filename-extension path) "")))
  (princ (strcat "\nجداکننده: " (cond ((= delim "\t") "Tab") ((= delim " ") "Space") (T delim))
                 "    سرستون: " (if hdr "دارد" "ندارد")))
  (princ (strcat "\nشمارهٔ ستون‌ها: نام=" (das-imp-colstr (nth 0 m))
                 "  X=" (das-imp-colstr (nth 1 m))
                 "  Y=" (das-imp-colstr (nth 2 m))
                 "  Z=" (das-imp-colstr (nth 3 m))
                 "  کد=" (das-imp-colstr (nth 4 m))))
  (if swapped
    (princ "\nستون X شمالی و ستون Y شرقی بود؛ جابه‌جا خوانده شدند."))
  (princ (strcat "\nتعداد نقاط خوانده‌شده: " (itoa (length recs))))
  (if bad
    (princ (strcat "\nردیف ناخوانا (مختصات عددی نیست): " (itoa (length bad))
                   "  ← خط‌ها: " (das-imp-join (das-imp-take bad 8) ", ")
                   (if (> (length bad) 8) ", ..." ""))))
  (foreach cfg *das-imp-cats*
    (setq n (das-imp-count-cat recs (car cfg)))
    (if (> n 0)
      (princ (strcat "\n  " (car cfg) "  " (nth 3 cfg) ": " (itoa n)))))
  (setq unc (das-imp-count-cat recs nil))
  (if (> unc 0)
    (princ (strcat "\n  بدون پیشوند/کد شناخته‌شده: " (itoa unc))))
  (setq n (das-imp-count-dups (mapcar 'car recs)))
  (if (> n 0) (princ (strcat "\nهشدار: " (itoa n) " نام تکراری")))
  (setq n (das-imp-count-dups
            (mapcar '(lambda (r) (strcat (rtos (nth 1 r) 2 2) "," (rtos (nth 2 r) 2 2))) recs)))
  (if (> n 0) (princ (strcat "\nهشدار: " (itoa n) " نقطه با مختصات تکراری")))
  (setq outR (length (vl-remove-if-not
                       '(lambda (r) (not (and (das-imp-in-e-p (nth 1 r)) (das-imp-in-n-p (nth 2 r)))))
                       recs)))
  (if (> outR 0)
    (princ (strcat "\nهشدار: " (itoa outR) " نقطه بیرون از محدودهٔ UTM ایران است")))
  (princ))

;; On-screen preview with grdraw: a coloured cross per point, nothing is
;; added to the drawing, so there is nothing to clean up on Esc/No.
(defun das-imp-show (recs / ext s pad r p col)
  (setq ext (das-imp-extent recs)
        s   (max (- (nth 2 ext) (nth 0 ext)) (- (nth 3 ext) (nth 1 ext))))
  (if (< s 1.0) (setq s 1.0))
  (setq pad (* 0.08 s))
  (vl-catch-all-apply 'vla-ZoomWindow
    (list (vlax-get-acad-object)
          (vlax-3d-point (- (nth 0 ext) pad) (- (nth 1 ext) pad) 0.0)
          (vlax-3d-point (+ (nth 2 ext) pad) (+ (nth 3 ext) pad) 0.0)))
  (redraw)
  (setq s (/ s 100.0))
  (foreach r recs
    (setq col (das-imp-cat-color (nth 6 r))
          p   (trans (list (nth 1 r) (nth 2 r) 0.0) 0 1))
    (grdraw (list (- (car p) s) (cadr p)) (list (+ (car p) s) (cadr p)) col)
    (grdraw (list (car p) (- (cadr p) s)) (list (car p) (+ (cadr p) s)) col))
  (princ))

(defun das-imp-ask-map (m ncol / out k v cur lbl)
  (princ (strcat "\nستون‌ها از ۱ تا " (itoa ncol) " شماره‌گذاری می‌شوند؛ ۰ یعنی «ندارد»."))
  (setq out '() k 0)
  (foreach lbl '("نام نقطه" "X (شرقی)" "Y (شمالی)" "Z (ارتفاع)" "کد")
    (setq cur (nth k m)
          cur (if cur (1+ cur) 0)
          v   (getint (strcat "\nشمارهٔ ستون " lbl " <" (itoa cur) ">: ")))
    (if (null v) (setq v cur))
    (setq out (append out (list (if (and (> v 0) (<= v ncol)) (1- v) nil)))
          k   (1+ k)))
  (if (and (nth 1 out) (nth 2 out))
    out
    (progn (princ "\nستون X و Y الزامی است.") nil)))

;; Points that matched no prefix and no code: one question for all of them.
(defun das-imp-assign-rest (recs / n ans)
  (setq n (das-imp-count-cat recs nil))
  (if (> n 0)
    (progn
      (initget "B CP BL W R S C P X None")
      (setq ans (getkword (strcat "\n" (itoa n)
                  " نقطه پیشوند یا کد شناخته‌شده ندارد. همه به کدام دسته بروند؟ [B/CP/BL/W/R/S/C/P/X/None] <None>: ")))
      (if (and ans (/= ans "None"))
        (setq recs (mapcar '(lambda (r) (if (nth 6 r) r (das-imp-put r 6 ans))) recs)))))
  recs)

;; Draw: one POINT on the category layer, the name as TEXT on "<layer>-T".
;; Rejected points (X) are skipped unless includeX.
(defun das-imp-draw (recs includeX h / n r cat cfg lay p)
  (setq n 0)
  (foreach r recs
    (setq cat (nth 6 r))
    (if (and cat (or includeX (/= cat "X")))
      (progn
        (setq cfg (assoc cat *das-imp-cats*)
              lay (cadr cfg)
              p   (list (nth 1 r) (nth 2 r) (nth 3 r)))
        (entmakex (list '(0 . "POINT") (cons 8 lay) (cons 10 p)))
        (entmakex (list '(0 . "TEXT") (cons 8 (strcat lay "-T")) (cons 7 "DAS_ROMANS")
                        (cons 10 (list (+ (car p) (* 0.6 h)) (+ (cadr p) (* 0.6 h)) (caddr p)))
                        (cons 40 h) (cons 1 (car r)) (cons 62 256)))
        (setq n (1+ n)))))
  n)

;; Show the report + preview and ask. Returns (status recs include-X)
;; with status 'yes or 'no. "Map" re-asks the column numbers and loops.
(defun das-imp-confirm (path delim hdr rows ncol m sw / done incx built recs bad ans m2)
  (setq done nil incx nil recs nil)
  (while (not done)
    (setq built (das-imp-build rows m)
          recs  (das-imp-assign (car built))
          bad   (cadr built))
    (if (null recs)
      (progn
        (princ "\nهیچ ردیف معتبری با این ستون‌ها پیدا نشد.")
        (initget "Map No")
        (setq ans (getkword "\nستون‌ها را دستی تعیین کنید؟ [Map/No] <No>: "))
        (if (equal ans "Map")
          (progn
            (setq m2 (das-imp-ask-map m ncol))
            (if m2 (setq m m2 sw nil)))
          (setq done 'no)))
      (progn
        (das-imp-report path delim hdr m sw recs bad)
        (das-imp-show recs)
        (princ "\nYes = ترسیم   No = انصراف   Map = تعیین دستی ستون‌ها   IncludeX = نقاط مردود (X) هم ترسیم شوند")
        (initget "Yes No Map IncludeX")
        (setq ans (getkword "\nترسیم شود؟ [Yes/No/Map/IncludeX] <Yes>: "))
        (cond
          ((or (null ans) (equal ans "Yes")) (setq done 'yes))
          ((equal ans "IncludeX") (setq incx T done 'yes))
          ((equal ans "No") (setq done 'no))
          (T
           (setq m2 (das-imp-ask-map m ncol))
           (if m2 (setq m m2 sw nil)))))))
  (redraw)
  (list done recs incx))

;; Ask for the leftovers and the text height, then draw inside one undo mark.
(defun das-imp-commit (recs incx / th n cfg nx nu)
  (setq recs (das-imp-assign-rest recs))
  (setq th (getreal (strcat "\nارتفاع متن نام نقطه <" (rtos *das-imp-texth* 2 2) ">: ")))
  (if (and th (> th 0.0)) (setq *das-imp-texth* th))
  (vla-StartUndoMark (vla-get-ActiveDocument (vlax-get-acad-object)))
  (das-utils-ensure-style "DAS_ROMANS" "romans.shx")
  (foreach cfg *das-imp-cats*
    (if (> (das-imp-count-cat recs (car cfg)) 0)
      (progn
        (das-utils-ensure-layer (cadr cfg) (caddr cfg))
        (das-utils-ensure-layer (strcat (cadr cfg) "-T") (caddr cfg)))))
  (if (= (getvar "PDMODE") 0)
    (progn
      (setvar "PDMODE" 35)
      (princ "\nPDMODE روی ۳۵ تنظیم شد تا نقطه‌ها دیده شوند.")))
  (setq n  (das-imp-draw recs incx *das-imp-texth*)
        nx (das-imp-count-cat recs "X")
        nu (das-imp-count-cat recs nil))
  (princ (strcat "\nDASIMP: " (itoa n) " نقطه با نام ترسیم شد."))
  (if (and (not incx) (> nx 0))
    (princ (strcat "\n" (itoa nx) " نقطهٔ مردود (X) ترسیم نشد.")))
  (if (> nu 0)
    (princ (strcat "\n" (itoa nu) " نقطهٔ بدون دسته ترسیم نشد.")))
  (princ))

(defun c:dasimp (/ *error* das-imp-restore oldCmd path data delim hdr rows ncol m sw res)
  (setq oldCmd (getvar 'cmdecho))

  (defun das-imp-restore ()
    (setvar 'cmdecho oldCmd)
    (vl-catch-all-apply 'vla-EndUndoMark
      (list (vla-get-ActiveDocument (vlax-get-acad-object))))
    (princ))

  (defun *error* (msg)
    (das-imp-restore)
    (redraw)
    (if (not (wcmatch (strcase msg) "*BREAK*,*CANCEL*,*EXIT*,*QUIT*"))
      (princ (strcat "\nError: " msg)))
    (princ))

  (setvar 'cmdecho 0)
  (das-init-globals)
  (setq path (if (das-farsi-ok-p)
               (getfiled "انتخاب فایل نقاط (CSV یا TXT)" *das-imp-last-path* "csv;txt;prn;xyz" 0)
               nil))

  (cond
    ;; [FIX-25] DASIMP's messages are Farsi only, so it needs Farsi text support
    ((not (das-farsi-ok-p))
     (princ (strcat "\nDASIMP is not available here: its messages are in Farsi and "
                    (das-farsi-why) ". See the guide.")))
    ((null path) (princ "\nلغو شد."))
    (T
     (setq *das-imp-last-path* path
           data (das-imp-load path))
     (cond
       ((null data)
        (alert "فایل خالی است یا جداکنندهٔ ستون‌ها (ویرگول، نقطه‌ویرگول، Tab یا فاصله) تشخیص داده نشد.\nاگر فایل Unicode (UTF-16) است، آن را در Notepad با UTF-8 یا ANSI دوباره ذخیره کنید."))
       ((null (caddr data))
        (alert "فایل ردیف داده ندارد."))
       (T
        (setq delim (car data) hdr (cadr data) rows (caddr data)
              ncol  (length (cdr (car rows))))
        (setq m (das-imp-auto-map hdr rows))
        (if (null m)
          (progn
            (princ "\nستون‌های X و Y خودکار پیدا نشد؛ شمارهٔ ستون‌ها را وارد کنید.")
            (setq m (das-imp-ask-map (list nil 0 1 nil nil) ncol))))
        (if (null m)
          (princ "\nDASIMP: لغو شد.")
          (progn
            (setq sw (das-imp-maybe-swap rows m)
                  m  (car sw)
                  sw (cadr sw))
            (setq res (das-imp-confirm path delim hdr rows ncol m sw))
            (if (and (eq (car res) 'yes) (cadr res))
              (das-imp-commit (cadr res) (caddr res))
              (progn
                (vl-catch-all-apply 'vla-ZoomPrevious (list (vlax-get-acad-object)))
                (princ "\nDASIMP: چیزی ترسیم نشد.")))))))))

  (das-imp-restore)
  (princ))

(defun c:das  () (das-core nil))
(defun c:dasl () (das-core T))

(defun c:dasreset ()
  (setq *das-last-scale* nil *das-last-start* nil *das-last-offset* nil
        *das-last-mask* nil *das-last-mkstyle* nil *das-last-mksize* nil
        *das-last-small-th* nil *das-last-small-fac* nil *das-last-col-count* nil
        *das-last-prec-c* nil *das-last-prec-l* nil *das-last-prec-a* nil
        *das-last-nb-scale* nil *das-last-nb-offset* nil *das-last-withlen* nil
        *das-lang* nil *das-start-rule* nil *das-corner-max-angle* nil
        *das-applicant* nil *das-nat-id* nil *das-b-area* nil *das-zone* nil
        *das-imp-texth* nil *das-imp-last-path* nil)
  (das-init-globals)
  (princ "\n*** DAS settings reset. ***")
  (princ))

;; [FIX-1] 9.6 declared oldOsm / oldClay / oldCmd but never assigned
;; them, so the final (setvar 'clayer nil) threw "bad argument type"
;; on every single run of this command.
(defun c:dasxy (/ *error* das-xy-restore p textloc x y oldOsm oldClay oldCmd)
  (setq oldOsm  (getvar 'osmode)
        oldClay (getvar 'clayer)
        oldCmd  (getvar 'cmdecho))

  (defun das-xy-restore ()
    (setvar 'osmode oldOsm) (setvar 'clayer oldClay) (setvar 'cmdecho oldCmd)
    (vl-catch-all-apply 'vla-EndUndoMark
      (list (vla-get-ActiveDocument (vlax-get-acad-object))))
    (princ))

  (defun *error* (msg)
    (das-xy-restore)
    (if (not (wcmatch (strcase msg) "*BREAK*,*CANCEL*,*EXIT*,*QUIT*"))
      (princ (strcat "\nError: " msg)))
    (princ))

  (setvar 'cmdecho 0)
  (vla-StartUndoMark (vla-get-ActiveDocument (vlax-get-acad-object)))
  (das-init-globals)
  (das-utils-ensure-style "DAS_ROMANS" "romans.shx")
  (das-utils-ensure-layer "L-LEADER" 6)
  (setvar 'clayer "L-LEADER")
  (princ "\n--- Manual Coordinate Labeling (DASXY) ---")

  (while (setq p (getpoint (das-tr "\nنقطه مورد نظر را کلیک کنید (Esc برای خروج): "
                                  "\nClick the point (Esc to exit): ")))
    (setvar 'osmode 0)
    (if (setq textloc (getpoint p (das-tr "\nمحل درج متن: " "\nText location: ")))
      (progn
        (setq x (strcat "X=" (rtos (car p)  2 *das-last-prec-c*))
              y (strcat "Y=" (rtos (cadr p) 2 *das-last-prec-c*)))
        (command "_.LEADER" "_non" p "_non" textloc "" x y "")))
    (setvar 'osmode oldOsm))

  (das-xy-restore)
  (princ "\nDone.")
  (princ))

(princ (strcat "\nDAS Tools V" *das-version* " loaded.  Commands: DAS / DASL / DASUPDATE / DASQC / DASCHECK / DASIMP / DASXY / DASRESET"))
(princ "\nDAS Tools is free. Support its development: https://hamibash.com/dastools")
(if (not (das-farsi-ok-p))
  (princ (strcat "\nDAS: Farsi text cannot be shown here (" (das-farsi-why) "). The interface is in English. See the guide.")))
(princ)
