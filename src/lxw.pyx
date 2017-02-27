
cimport lxw

cdef class Workbook:
    cdef:
        lxw.lxw_workbook* _c_workbook

    def __cinit__(self, const char filename):
        self._c_workbook = lxw.workbook_new(&filename)

    def __dealloc__(self):
        if self._c_workbook is not NULL:
            lxw.lxw_workbook_free(self._c_workbook)
