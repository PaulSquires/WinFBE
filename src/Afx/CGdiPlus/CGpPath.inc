' ########################################################################################
' Microsoft Windows
' File: CGpPath.inc
' Contents: GDI+ Path classes.
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ########################################################################################
' CGpGraphicsPath Class
' ########################################################################################

' ========================================================================================
' * Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpGraphicsPath (BYVAL nFillMode AS FillMode = FillModeAlternate)
   m_Status = GdipCreatePath(nFillMode, @m_pPath)
END CONSTRUCTOR
' ========================================================================================

' =====================================================================================
' Creates a GraphicsPath object based on an array of points, an array of types, and a fill mode.
' Default value for filolMode: FillModeAlternate(0).
' =====================================================================================
PRIVATE CONSTRUCTOR CGpGraphicsPath (BYVAL pts AS GpPointF PTR, BYVAL types AS BYTE PTR, BYVAL nCount AS LONG, BYVAL nFillMode AS FillMode = FillModeAlternate)
   m_Status = GdipCreatePath2(pts, types, nCount, nFillmode, @m_pPath)
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR CGpGraphicsPath (BYVAL pts AS GpPoint PTR, BYVAL types AS BYTE PTR, BYVAL nCount AS LONG, BYVAL nFillMode AS FillMode = FillModeAlternate)
   m_Status = GdipCreatePath2I(pts, types, nCount, nFillmode, @m_pPath)
END CONSTRUCTOR
' =====================================================================================

' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR CGpGraphicsPath
   IF m_pPath THEN GdipDeletePath(m_pPath)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpGraphicsPath.GetLastStatus () AS GpStatus
   RETURN m_Status
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpGraphicsPath.SetStatus (BYVAL status AS GpStatus) AS GpStatus
   m_Status = status
   RETURN m_Status
END FUNCTION
' ========================================================================================

' =====================================================================================
' Creates a new Path object based on this Path object.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.Clone (BYVAL pCloneGraphicsPath AS CGpGraphicsPath PTR) AS GpStatus
   IF pCloneGraphicsPath = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pCloneGraphicsPath->m_pPath THEN GdipDeletePath(pCloneGraphicsPath->m_pPath)
   RETURN SetStatus(GdipClonePath(m_pPath, @pCloneGraphicsPath->m_pPath))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Empties the path and sets the fill mode to FillModeAlternate.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.Reset () AS GpStatus
   RETURN SetStatus(GdipResetPath(m_pPath))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the fill mode of this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.GetFillMode () AS FillMode
   DIM nFillmode AS FillMode = FillModeAlternate
   m_Status = GdipGetPathFillMode(m_pPath, @nFillmode)
   RETURN nFillmode
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the fill mode of this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.SetFillMode (BYVAL nFillmode AS FillMode) AS GpStatus
   RETURN SetStatus(GdipSetPathFillMode(m_pPath, nFillmode))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets an array of points and an array of point types from this path. Together, these
' two arrays define the lines, curves, figures, and markers of this path.
' Parameter declared as ANY PTR because conflicts with the FB declares regarding GpPathData.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.GetPathData (BYVAL pPathData AS ANY PTR) AS GpStatus
   RETURN SetStatus(GdipGetPathData(m_pPath, pPathData))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Starts a new figure without closing the current figure. Subsequent points added to
' this path are added to the new figure.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.StartFigure () AS GpStatus
   RETURN SetStatus(GdipStartPathFigure(m_pPath))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Closes the current figure of this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.CloseFigure () AS GpStatus
   RETURN SetStatus(GdipClosePathFigure(m_pPath))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Closes all open figures in this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.CloseAllFigures () AS GpStatus
   RETURN SetStatus(GdipClosePathFigures(m_pPath))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Designates the last point in this path as a marker point.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.SetMarker () AS GpStatus
   RETURN SetStatus(GdipSetPathMarker(m_pPath))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Clears the markers from this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.ClearMarkers () AS GpStatus
   RETURN SetStatus(GdipClearPathMarkers(m_pPath))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Reverses the order of the points that define this path's lines and curves.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.Reverse () AS GpStatus
   RETURN SetStatus(GdipReversePath(m_pPath))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the ending point of the last figure in this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.GetLastPoint (BYVAL lastPoint AS GpPointF PTR) AS GpStatus
   RETURN SetStatus(GdipGetPathLastPoint(m_pPath, lastPoint))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a line to the current figure of this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddLine (BYVAL x1 AS SINGLE, BYVAL y1 AS SINGLE, BYVAL x2 AS SINGLE, BYVAL y2 AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipAddPathLine(m_pPath, x1, y1, x2, y1))
END FUNCTION
' =====================================================================================
' =====================================================================================
' * Adds a line to the current figure of this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddLine (BYVAL x1 AS INT_, BYVAL y1 AS INT_, BYVAL x2 AS INT_, BYVAL y2 AS INT_) AS GpStatus
   RETURN SetStatus(GdipAddPathLineI(m_pPath, x1, y1, x2, y1))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddLine (BYVAL pt1 AS GpPointF PTR, BYVAL pt2 AS GpPointF PTR) AS GpStatus
   RETURN SetStatus(GdipAddPathLine(m_pPath, pt1->x, pt1->y, pt2->x, pt2->y))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddLine (BYVAL pt1 AS GpPoint PTR, BYVAL pt2 AS GpPoint PTR) AS GpStatus
   RETURN SetStatus(GdipAddPathLineI(m_pPath, pt1->x, pt1->y, pt2->x, pt2->y))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a sequence of connected lines to the current figure of this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddLines (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipAddPathLine2(m_pPath, pts, nCount))
END FUNCTION
' =====================================================================================
' =====================================================================================
' * Adds a sequence of connected lines to the current figure of this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddLines (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipAddPathLine2I(m_pPath, pts, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds an elliptical arc to the current figure of this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddArc (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, _
   BYVAL nHeight AS SINGLE, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipAddPathArc(m_pPath, x, y, nWidth, nHeight, startAngle, sweepAngle))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddArc (BYVAL rcf AS GpRectF PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipAddPathArc(m_pPath, rcf->x, rcf->y, rcf->Width, rcf->Height, startAngle, sweepAngle))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddArc (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, _
   BYVAL nHeight AS INT_, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipAddPathArcI(m_pPath, x, y, nWidth, nHeight, startAngle, sweepAngle))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddArc (BYVAL rc AS GpRect PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipAddPathArcI(m_pPath, rc->x, rc->y, rc->Width, rc->Height, startAngle, sweepAngle))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a Bézier spline to the current figure of this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddBezier (BYVAL x1 AS SINGLE, BYVAL y1 AS SINGLE, BYVAL x2 AS SINGLE, BYVAL y2 AS SINGLE, _
   BYVAL x3 AS SINGLE, BYVAL y3 AS SINGLE, BYVAL x4 AS SINGLE, BYVAL y4 AS SINGLE) AS GpStatus
   m_Status = GdipAddPathBezier(m_pPath, x1, y1, x2, y2, x3, y3, x4, y4)
   FUNCTION = m_Status
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddBezier (BYVAL x1 AS INT_, BYVAL y1 AS INT_, BYVAL x2 AS INT_, BYVAL y2 AS INT_, _
   BYVAL x3 AS INT_, BYVAL y3 AS INT_, BYVAL x4 AS INT_, BYVAL y4 AS INT_) AS GpStatus
   m_Status = GdipAddPathBezierI(m_pPath, x1, y1, x2, y2, x3, y3, x4, y4)
   FUNCTION = m_Status
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddBezier (BYVAL pt1 AS GpPointF PTR, BYVAL pt2 AS GpPointF PTR, BYVAL pt3 AS GpPointF PTR, BYVAL pt4 AS GpPointF PTR) AS GpStatus
   RETURN SetStatus(GdipAddPathBezier(m_pPath, pt1->x, pt1->y, pt2->x, pt2->y, pt3->x, pt3->y, pt4->x, pt4->y))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddBezier (BYVAL pt1 AS GpPoint PTR, BYVAL pt2 AS GpPoint PTR, BYVAL pt3 AS GpPoint PTR, BYVAL pt4 AS GpPoint PTR) AS GpStatus
   RETURN SetStatus(GdipAddPathBezierI(m_pPath, pt1->x, pt1->y, pt2->x, pt2->y, pt3->x, pt3->y, pt4->x, pt4->y))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a sequence of connected Bézier splines to the current figure of this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddBeziers (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipAddPathBeziers(m_pPath, pts, nCount))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddBeziers (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipAddPathBeziersI(m_pPath, pts, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a cardinal spline to the current figure of this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddCurve (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipAddPathCurve(m_pPath, pts, nCount))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddCurve (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipAddPathCurveI(m_pPath, pts, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a cardinal spline to the current figure of this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddCurve (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipAddPathCurve2(m_pPath, pts, nCount, tension))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddCurve (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipAddPathCurve2I(m_pPath, pts, nCount, tension))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a cardinal spline to the current figure of this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddCurve (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_, BYVAL offset AS INT_, _
   BYVAL numberOfSegments AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipAddPathCurve3(m_pPath, pts, nCount, offset, numberOfSegments, tension))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddCurve (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_, BYVAL offset AS INT_, _
   BYVAL numberOfSegments AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipAddPathCurve3I(m_pPath, pts, nCount, offset, numberOfSegments, tension))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a closed cardinal spline to this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddClosedCurve (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipAddPathClosedCurve(m_pPath, pts, nCount))
END FUNCTION
' =====================================================================================
' =====================================================================================
' Adds a closed cardinal spline to this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddClosedCurve (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipAddPathClosedCurve2(m_pPath, pts, nCount, tension))
END FUNCTION
' =====================================================================================
' =====================================================================================
' * Adds a closed cardinal spline to this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddClosedCurve (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipAddPathClosedCurveI(m_pPath, pts, nCount))
END FUNCTION
' =====================================================================================
' =====================================================================================
' Adds a closed cardinal spline to this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddClosedCurve (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipAddPathClosedCurve2I(m_pPath, pts, nCount, tension))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a rectangle to this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddRectangle (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipAddPathRectangle(m_pPath, x, y, nWidth, nHeight))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddRectangle (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
   RETURN SetStatus(GdipAddPathRectangleI(m_pPath, x, y, nWidth, nHeight))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddRectangle (BYVAL rcf AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipAddPathRectangle(m_pPath, rcf->x, rcf->y, rcf->Width, rcf->Height))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddRectangle (BYVAL rcf AS GpRect PTR) AS GpStatus
   RETURN SetStatus(GdipAddPathRectangleI(m_pPath, rcf->x, rcf->y, rcf->Width, rcf->Height))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a sequence of rectangles to this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddRectangles (BYVAL rects AS GpRectF PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipAddPathRectangles(m_pPath, rects, nCount))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddRectangles (BYVAL rects AS GpRect PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipAddPathRectanglesI(m_pPath, rects, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds an ellipse to this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddEllipse (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipAddPathEllipse(m_pPath, x, y, nWidth, nHeight))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddEllipse (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
   RETURN SetStatus(GdipAddPathEllipseI(m_pPath, x, y, nWidth, nHeight))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddEllipse (BYVAL rcf AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipAddPathEllipse(m_pPath, rcf->x, rcf->y, rcf->Width, rcf->Height))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddEllipse (BYVAL rc AS GpRect PTR) AS GpStatus
   RETURN SetStatus(GdipAddPathEllipseI(m_pPath, rc->x, rc->y, rc->Width, rc->Height))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a pie to this path. An arc is a portion of an ellipse, and a pie is a portion of
' the area enclosed by an ellipse. A pie is bounded by an arc and two lines (edges) that
' go from the center of the ellipse to the endpoints of the arc.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddPie (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, _
   BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipAddPathPie(m_pPath, x, y, nWidth, nHeight, startAngle, sweepAngle))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddPie (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, _
   BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipAddPathPieI(m_pPath, x, y, nWidth, nHeight, startAngle, sweepAngle))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddPie (BYVAL rcf AS GpRectF PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipAddPathPie(m_pPath, rcf->x, rcf->y, rcf->Width, rcf->Height, startAngle, sweepAngle))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddPie (BYVAL rc AS GpRect PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipAddPathPieI(m_pPath, rc->x, rc->y, rc->Width, rc->Height, startAngle, sweepAngle))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a polygon to this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddPolygon (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipAddPathPolygon(m_pPath, pts, nCount))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddPolygon (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipAddPathPolygonI(m_pPath, pts, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a path to this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddPath (BYVAL pAddingPath AS CGpGraphicsPath PTR, BYVAL bConnect AS BOOL) AS GpStatus
   RETURN SetStatus(GdipAddPathPath(m_pPath, pAddingPath->m_pPath, bConnect))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds the outlines of a string to this path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddString (BYVAL pString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFamily AS CGpFontFamily PTR, BYVAL nStyle AS INT_, _
   BYVAL emSize AS SINGLE, BYVAL layoutRect AS GpRectF PTR, BYVAL pFormat AS CGpStringFormat PTR) AS GpStatus
   RETURN SetStatus(GdipAddPathString(m_pPath, pString, length, IIF(pFamily, pFamily->m_pFontFamily, NULL), _
      nStyle, emSize, layoutRect, IIF(pFormat, pFormat->m_pStringFormat, NULL)))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddString (BYVAL pString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFamily AS CGpFontFamily PTR, BYVAL nStyle AS INT_, _
   BYVAL emSize AS SINGLE, BYVAL layoutRect AS GpRect PTR, BYVAL pFormat AS CGpStringFormat PTR) AS GpStatus
   RETURN SetStatus(GdipAddPathStringI(m_pPath, pString, length, IIF(pFamily, pFamily->m_pFontFamily, NULL), _
      nStyle, emSize, layoutRect, IIF(pFormat, pFormat->m_pStringFormat, NULL)))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddString (BYVAL pString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFamily AS CGpFontFamily PTR, BYVAL nStyle AS INT_, _
   BYVAL emSize AS SINGLE, BYVAL origin AS GpPointF PTR, BYVAL pFormat AS CGpStringFormat PTR) AS GpStatus
   DIM layoutRect AS GpRectF = TYPE<GpRectF>(origin->x, origin->y, 0.0, 0.0)
   RETURN SetStatus(GdipAddPathString(m_pPath, pString, length, IIF(pFamily, pFamily->m_pFontFamily, NULL), _
      nStyle, emSize, @layoutRect, IIF(pFormat, pFormat->m_pStringFormat, NULL)))
END FUNCTION
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.AddString (BYVAL pString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFamily AS CGpFontFamily PTR, BYVAL nStyle AS INT_, _
   BYVAL emSize AS SINGLE, BYVAL origin AS GpPoint PTR, BYVAL pFormat AS CGpStringFormat PTR) AS GpStatus
   DIM layoutRect AS GpRect = TYPE<GpRect>(origin->x, origin->y, 0, 0)
   RETURN SetStatus(GdipAddPathStringI(m_pPath, pString, length, IIF(pFamily, pFamily->m_pFontFamily, NULL), _
      nStyle, emSize, @layoutRect, IIF(pFormat, pFormat->m_pStringFormat, NULL)))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Multiplies each of this path's data points by a specified matrix.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.TransForm (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   RETURN SetStatus(GdipTransformPath(m_pPath, pMatrix->m_pMatrix))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets a bounding rectangle for this path
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.GetBounds (BYVAL bounds AS GpRectF PTR, BYVAL pMatrix AS CGpMatrix PTR, BYVAL pPen AS CGpPen PTR) AS GpStatus
   RETURN SetStatus(GdipGetPathWorldBounds(m_pPath, bounds, IIF(pMatrix, pMatrix->m_pMatrix, NULL), IIF(pPen, pPen->m_pPen, NULL)))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.GetBounds (BYVAL bounds AS GpRect PTR, BYVAL pMatrix AS CGpMatrix PTR, BYVAL pPen AS CGpPen PTR) AS GpStatus
   RETURN SetStatus(GdipGetPathWorldBoundsI(m_pPath, bounds, IIF(pMatrix, pMatrix->m_pMatrix, NULL), IIF(pPen, pPen->m_pPen, NULL)))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Applies a transformation to this path and converts each curve in the path to a
' sequence of connected lines.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.Flatten (BYVAL pMatrix AS CGpMatrix PTR = NULL, BYVAL flatness AS SINGLE = FlatnessDefault) AS GpStatus
   RETURN SetStatus(GdipFlattenPath(m_pPath, IIF(pMatrix, pMatrix->m_pMatrix, NULL), flatness))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Replaces this path with curves that enclose the area that is filled when this path is
' drawn by a specified pen. The GraphicsPath::Widen FUNCTION also flattens the path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.Widen (BYVAL pPen AS CGpPen PTR, BYVAL pMatrix AS CGpMatrix PTR = NULL, BYVAL flatness AS SINGLE = FlatnessDefault) AS GpStatus
   RETURN SetStatus(GdipWidenPath(m_pPath, pPen->m_pPen, IIF(pMatrix, pMatrix->m_pMatrix, NULL), flatness))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Transforms and flattens this path, and then converts this path's data points so that
' they represent only the outline of the path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.Outline (BYVAL pMatrix AS CGpMatrix PTR = NULL, BYVAL flatness AS SINGLE = FlatnessDefault) AS GpStatus
   RETURN SetStatus(GdipWindingModeOutline(m_pPath, IIF(pMatrix, pMatrix->m_pMatrix, NULL), flatness))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Applies a warp transformation to this path. Warp also flattens (converts to a sequence
' of straight lines) the path. Default warpMode value: WarpModePerspective (0).
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.Warp (BYVAL destPoints AS GpPointF PTR, BYVAL nCount AS INT_, BYVAL srcRect AS GpRectF PTR, _
   BYVAL pMatrix AS CGpMatrix PTR = NULL, BYVAL nWarpMode AS WarpMode = WarpModePerspective, BYVAL Flatness AS SINGLE = FlatnessDefault) AS GpStatus
   RETURN SetStatus(GdipWarpPath(m_pPath, IIF(pMatrix, pMatrix->m_pMatrix, NULL), destPoints, nCount, _
          srcRect->x, srcRect->y, srcRect->width, srcRect->height, nWarpMode, flatness))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the number of points in this path's array of data points. This is the same as
' the number of types in the path's array of point types.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.GetPointCount () AS INT_
   DIM nCount AS INT_
   m_Status = GdipGetPointCount(m_pPath, @nCount)
   RETURN nCount
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets this path's array of point types.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.GetPathTypes (BYVAL types AS BYTE PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipGetPathTypes(m_pPath, types, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets this path's array of points. The array contains the endpoints and control points
' of the lines and Bézier splines that are used to draw the path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.GetPathPoints (BYREF pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipGetPathPoints(m_pPath, pts, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Determines whether a specified point lies in the area that is filled when this path
' is filled by a specified Graphics object.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.IsVisible (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL pGraphics AS CGpGraphics PTR = NULL) AS BOOLEAN
   DIM bResult AS BOOL
   m_Status = GdipIsVisiblePathPoint(m_pPath, x, y, IIF(pGraphics, pGraphics->m_pGraphics, NULL), @bResult)
   RETURN bResult
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.IsVisible (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL pGraphics AS CGpGraphics PTR = NULL) AS BOOLEAN
   DIM bResult AS BOOL
   m_Status = GdipIsVisiblePathPointI(m_pPath, x, y, IIF(pGraphics, pGraphics->m_pGraphics, NULL), @bResult)
   RETURN bResult
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.IsVisible (BYVAL pt AS GpPointF PTR, BYVAL pGraphics AS CGpGraphics PTR = NULL) AS BOOLEAN
   DIM bResult AS BOOL
   m_Status = GdipIsVisiblePathPoint(m_pPath, pt->x, pt->y, IIF(pGraphics, pGraphics->m_pGraphics, NULL), @bResult)
   RETURN bResult
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.IsVisible (BYVAL pt AS GpPoint PTR, BYVAL pGraphics AS CGpGraphics PTR = NULL) AS BOOLEAN
   DIM bResult AS BOOL
   m_Status = GdipIsVisiblePathPointI(m_pPath, pt->x, pt->y, IIF(pGraphics, pGraphics->m_pGraphics, NULL), @bResult)
   RETURN bResult
END FUNCTION
' =====================================================================================

' =====================================================================================
' Determines whether a specified point touches the outline of this path when the path
' is drawn by a specified Graphics object and a specified pen.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.IsOutlineVisible (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL pPen AS CGpPen PTR, BYVAL pGraphics AS CGpGraphics PTR = NULL) AS BOOLEAN
   DIM bResult AS BOOL
   m_Status = GdipIsOutlineVisiblePathPoint(m_pPath, x, y, IIF(pPen, pPen->m_pPen, NULL), IIF(pGraphics, pGraphics->m_pGraphics, NULL), @bResult)
   FUNCTION = bResult
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.IsOutlineVisible (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL pPen AS CGpPen PTR, BYVAL pGraphics AS CGpGraphics PTR = NULL) AS BOOLEAN
   DIM bResult AS BOOL
   m_Status = GdipIsOutlineVisiblePathPointI(m_pPath, x, y, IIF(pPen, pPen->m_pPen, NULL), IIF(pGraphics, pGraphics->m_pGraphics, NULL), @bResult)
   FUNCTION = bResult
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.IsOutlineVisible (BYVAL pt AS GpPointF PTR, BYVAL pPen AS CGpPen PTR, BYVAL pGraphics AS CGpGraphics PTR = NULL) AS BOOLEAN
   DIM bResult AS BOOL
   m_Status = GdipIsOutlineVisiblePathPoint(m_pPath, pt->x, pt->y, IIF(pPen, pPen->m_pPen, NULL), IIF(pGraphics, pGraphics->m_pGraphics, NULL), @bResult)
   FUNCTION = bResult
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPath.IsOutlineVisible (BYVAL pt AS GpPoint PTR, BYVAL pPen AS CGpPen PTR, BYVAL pGraphics AS CGpGraphics PTR = NULL) AS BOOLEAN
   DIM bResult AS BOOL
   m_Status = GdipIsOutlineVisiblePathPointI(m_pPath, pt->x, pt->y, IIF(pPen, pPen->m_pPen, NULL), IIF(pGraphics, pGraphics->m_pGraphics, NULL), @bResult)
   FUNCTION = bResult
END FUNCTION
' =====================================================================================

' ########################################################################################

' ########################################################################################
' CGpGraphicsPathIterator Class
' ########################################################################################

' =====================================================================================
' Creates a new CGpGraphicsPathIterator object and associates it with a CGpGraphicsPath object.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpGraphicsPathIterator (BYVAL pPath AS CGpGraphicsPath PTR)
   m_Status = GdipCreatePathIter(@m_pPathIretator, pPath->m_pPath)
END CONSTRUCTOR
' =====================================================================================

' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR CGpGraphicsPathIterator
   IF m_pPathIretator THEN GdipDeletePathIter(m_pPathIretator)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpGraphicsPathIterator.GetLastStatus () AS GpStatus
   RETURN m_Status
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpGraphicsPathIterator.SetStatus (BYVAL status AS GpStatus) AS GpStatus
   m_Status = status
   RETURN m_Status
END FUNCTION
' ========================================================================================

' =====================================================================================
' Gets the starting index and the ending index of the next subpath (figure) in this
' iterator's associated path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPathIterator.NextSubpath (BYVAL startIndex AS INT_ PTR, BYVAL endIndex AS INT_ PTR, BYVAL isClosed AS BOOL PTR) AS INT_
   DIM resultCount AS INT_
   m_Status = GdipPathIterNextSubpath(m_pPathIretator, @resultCount, startIndex, endIndex, isClosed)
   RETURN resultCount
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the the next figure (subpath) from this iterator's associated path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPathIterator.NextSubpathPath (BYVAL pPath AS CGpGraphicsPath PTR, BYVAL isClosed AS BOOL PTR) AS INT_
   DIM resultCount AS INT_
   m_Status = GdipPathIterNextSubpathPath(m_pPathIretator, @resultCount, pPath->m_pPath, isClosed)
   RETURN resultCount
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the starting index and the ending index of the next group of data points that
' all have the same type.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPathIterator.NextPathType (BYVAL pathType AS BYTE PTR, BYVAL startIndex AS INT_ PTR, BYVAL endIndex AS INT_ PTR) AS INT_
   DIM resultCount AS INT_
   m_Status = GdipPathIterNextPathType(m_pPathIretator, @resultCount, pathType, startIndex, endIndex)
   RETURN resultCount
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the starting index and the ending index of the next marker-delimited section in
' this iterator's associated path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPathIterator.NextMarker (BYVAL startIndex AS INT_ PTR, BYVAL endIndex AS INT_ PTR) AS INT_
   DIM resultCount AS INT_
   m_Status = GdipPathIterNextMarker(m_pPathIretator, @resultCount, startIndex, endIndex)
   RETURN resultCount
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPathIterator.NextMarker (BYVAL pPath AS CGpGraphicsPath PTR) AS INT_
   DIM resultCount AS INT_
   m_Status = GdipPathIterNextMarkerPath(m_pPathIretator, @resultCount, pPath->m_pPath)
   RETURN resultCount
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the number of data points in the path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPathIterator.GetCount () AS INT_
   DIM resultCount AS INT_
   m_Status = GdipPathIterGetCount(m_pPathIretator, @resultCount)
   RETURN resultCount
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the number of subpaths (also called figures) in the path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPathIterator.GetSubpathCount () AS INT_
   DIM resultCount AS INT_
   m_Status = GdipPathIterGetSubpathCount(m_pPathIretator, @resultCount)
   RETURN resultCount
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the number of subpaths (also called figures) in the path.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPathIterator.HasCurve () AS BOOLEAN
   DIM bHasCurve AS BOOL
   m_Status = GdipPathIterHasCurve(m_pPathIretator, @bHasCurve)
   FUNCTION = bHasCurve
END FUNCTION
' =====================================================================================

' =====================================================================================
' Rewinds this iterator to the beginning of its associated path.
' =====================================================================================
PRIVATE SUB CGpGraphicsPathIterator.Rewind
   GdipPathIterRewind(m_pPathIretator)
END SUB
' =====================================================================================

' =====================================================================================
' Copies the path's data points to a PointF array and copies the path's point types to a BYTE array.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPathIterator.Enumerate (BYVAL pts AS GpPointF PTR, BYVAL types AS BYTE PTR, BYVAL nCount AS INT_) AS INT_
   DIM resultCount AS INT_
   m_Status = GdipPathIterEnumerate(m_pPathIretator, @resultCount, pts, types, nCount)
   RETURN resultCount
END FUNCTION
' =====================================================================================

' =====================================================================================
' Copies a subset of the path's data points to a PointF array and copies a subset of
' the path's point types to a BYTE array.
' =====================================================================================
PRIVATE FUNCTION CGpGraphicsPathIterator.CopyData (BYVAL pts AS GpPointF PTR, BYVAL types AS BYTE PTR, BYVAL startIndex AS INT_, BYVAL endIndex AS INT_) AS INT_
   DIM resultCount AS INT_
   m_Status = GdipPathIterCopyData(m_pPathIretator, @resultCount, pts, types, startIndex, endIndex)
   RETURN resultCount
END FUNCTION
' =====================================================================================

' ########################################################################################
