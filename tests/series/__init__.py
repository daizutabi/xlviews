if __name__ == "__main__":
    import xlwings as xw
    from xlwings.constants import ChartType, LineStyle, MarkerStyle

    from xlviews.axes import Axes
    from xlviews.common import quit_apps
    from xlviews.utils import rgb

    quit_apps()
    book = xw.Book()
    ax = Axes(left=200, chart_type=ChartType.xlXYScatter)
    s = ax.add_series(x=[1, 2, 3], y=[4, 5, 6])

    s.set(marker="o", line=None, color="red", alpha=0.1, size=5, weight=2)
    s.set(marker=None, line="-", color="red", alpha=0.1, weight=2)
    s.set(marker="o", line="-", color="red", alpha=0.1, size=5, weight=2)

    s.api.MarkerStyle = MarkerStyle.xlMarkerStyleCircle
    s.api.MarkerSize = 5
    s.api.Format.Fill.Visible = True
    s.api.Format.Fill.BackColor.RGB = rgb("yellow")
    s.api.Format.Fill.Transparency = 0.5
    s.api.Format.Fill.ForeColor.RGB = rgb("yellow")
    s.api.Border.LineStyle = LineStyle.xlContinuous
    s.api.Format.Line.Visible = True
    s.api.Format.Line.Weight = 4
    s.api.Format.Line.Transparency = 0.5 / 2
    s.api.Format.Line.ForeColor.RGB = rgb("red")
    s.api.Border.LineStyle = LineStyle.xlContinuous

    color = "red"
    alpha = 0.0
    size = 7
    s.api.MarkerStyle = MarkerStyle.xlMarkerStyleCircle
    s.api.MarkerStyle = MarkerStyle.xlMarkerStyleNone
    s.api.MarkerSize = size
    s.api.Format.Fill.Visible = True
    s.api.Format.Fill.BackColor.RGB = rgb(color)
    s.api.Format.Fill.Transparency = alpha
    s.api.Format.Fill.ForeColor.RGB = rgb(color)
    s.api.Border.LineStyle = LineStyle.xlContinuous
    s.api.Format.Line.Visible = True
    s.api.Format.Line.Weight = min(size / 4, 2)
    s.api.Format.Line.Transparency = alpha / 2
    s.api.Format.Line.ForeColor.RGB = rgb(color)
    # s.api.Border.LineStyle = LineStyle.xlLineStyleNone
    s.api.Border.LineStyle = LineStyle.xlDash
