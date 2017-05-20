# coding: utf-8
from libc.stdint cimport uint32_t, uint16_t, uint8_t

from impeller.c_workbook cimport WorkBook

cdef extern from "xlsxwriter.h":
    ctypedef enum lxw_chart_type:
        LXW_CHART_NONE = 0,
    
        # Area chart.
        LXW_CHART_AREA,
    
        # Area chart - stacked.
        LXW_CHART_AREA_STACKED,
    
        # Area chart - percentage stacked.
        LXW_CHART_AREA_STACKED_PERCENT,
    
        # Bar chart.
        LXW_CHART_BAR,
    
        # Bar chart - stacked.
        LXW_CHART_BAR_STACKED,
    
        # Bar chart - percentage stacked.
        LXW_CHART_BAR_STACKED_PERCENT,
    
        # Column chart.
        LXW_CHART_COLUMN,
    
        # Column chart - stacked.
        LXW_CHART_COLUMN_STACKED,
    
        # Column chart - percentage stacked.
        LXW_CHART_COLUMN_STACKED_PERCENT,
    
        # Doughnut chart.
        LXW_CHART_DOUGHNUT,
    
        # Line chart.
        LXW_CHART_LINE,
    
        # Pie chart.
        LXW_CHART_PIE,
    
        # Scatter chart.
        LXW_CHART_SCATTER,
    
        # Scatter chart - straight.
        LXW_CHART_SCATTER_STRAIGHT,
    
        # Scatter chart - straight with markers.
        LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS,
    
        # Scatter chart - smooth.
        LXW_CHART_SCATTER_SMOOTH,
    
        # Scatter chart - smooth with markers.
        LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS,
    
        # Radar chart.
        LXW_CHART_RADAR,
    
        # Radar chart - with markers.
        LXW_CHART_RADAR_WITH_MARKERS,
    
        # Radar chart - filled.
        LXW_CHART_RADAR_FILLED

    ctypedef struct lxw_chart:
        pass

    ctypedef struct lxw_chart_series:
        pass

    lxw_chart_series *chart_add_series(lxw_chart *chart, const char *categories, const char *values);

# TODO: Simulate subclasses with factory?
# Factory return proper type by chart type from lxw_chart_type
cdef class Chart:
    cdef lxw_chart* this_ptr;
    cdef int chart_type;

    cdef void _set_ptr(self, lxw_chart* ptr);

    cdef void _add_chart(self, WorkBook wb, uint8_t chart_type);

    # Return series object?
    cpdef void add_series(self, dict options=*);
