
from libc.stdio cimport *
from libc.stdint cimport *
from libc.errno cimport *

cdef extern from "../libxlsxwriter/include/xlsxwriter/common.h" nogil:
    ctypedef unsigned int lxw_row_t;

    ctypedef unsigned short lxw_col_t;

    enum lxw_boolean:
        pass

    ctypedef enum lxw_error:
        pass

    ctypedef struct lxw_datetime:
        pass

    enum lxw_custom_property_types:
        pass

    ctypedef struct lxw_tuple:
        pass

    ctypedef struct lxw_custom_property:
        pass


cdef extern from "../libxlsxwriter/include/xlsxwriter/workbook.h" nogil:
    ctypedef struct lxw_workbook:
        pass

    lxw_workbook *workbook_new(const char *filename);

    void lxw_workbook_free(lxw_workbook *workbook);


cdef extern from "../libxlsxwriter/src/workbook.c" nogil:
    lxw_workbook *workbook_new(const char *filename);

    void lxw_workbook_free(lxw_workbook *workbook);


cdef extern from "../libxlsxwriter/include/xlsxwriter/worksheet.h":
    ctypedef struct lxw_row_col_options:
        pass

    ctypedef struct lxw_col_options:
        pass

    ctypedef struct lxw_merged_range:
        pass

    ctypedef struct lxw_repeat_rows:
        pass

    ctypedef struct lxw_repeat_cols:
        pass

    ctypedef struct lxw_print_area:
        pass

    ctypedef struct lxw_autofilter:
        pass

    ctypedef struct lxw_panes:
        pass

    ctypedef struct lxw_selection:
        pass

    ctypedef struct lxw_image_options:
        pass

    ctypedef struct lxw_header_footer_options:
        pass

    ctypedef struct lxw_protection:
        pass

    ctypedef struct lxw_worksheet:
        pass

    ctypedef struct lxw_worksheet_init_data:
        pass

    ctypedef struct lxw_row:
        pass

    ctypedef struct lxw_cell:
        pass

    lxw_error worksheet_write_number(lxw_worksheet *worksheet,
                                     lxw_row_t row,
                                     lxw_col_t col,
                                     double number,
                                     lxw_format *format);

    lxw_error worksheet_write_string(lxw_worksheet *worksheet,
                                     lxw_row_t row,
                                     lxw_col_t col, const char *string,
                                     lxw_format *format);

    lxw_error worksheet_write_formula(lxw_worksheet *worksheet,
                                      lxw_row_t row,
                                      lxw_col_t col, const char *formula,
                                      lxw_format *format);

    lxw_error worksheet_write_array_formula(lxw_worksheet *worksheet,
                                            lxw_row_t first_row,
                                            lxw_col_t first_col,
                                            lxw_row_t last_row,
                                            lxw_col_t last_col,
                                            const char *formula,
                                            lxw_format *format);

    lxw_error worksheet_write_array_formula_num(lxw_worksheet *worksheet,
                                                lxw_row_t first_row,
                                                lxw_col_t first_col,
                                                lxw_row_t last_row,
                                                lxw_col_t last_col,
                                                const char *formula,
                                                lxw_format *format,
                                                double result);

    lxw_error worksheet_write_datetime(lxw_worksheet *worksheet,
                                       lxw_row_t row,
                                       lxw_col_t col, lxw_datetime *datetime,
                                       lxw_format *format);

    lxw_error worksheet_write_url_opt(lxw_worksheet *worksheet,
                                      lxw_row_t row_num,
                                      lxw_col_t col_num, const char *url,
                                      lxw_format *format, const char *string,
                                      const char *tooltip);

    lxw_error worksheet_write_url(lxw_worksheet *worksheet,
                                  lxw_row_t row,
                                  lxw_col_t col, const char *url,
                                  lxw_format *format);

    lxw_error worksheet_write_boolean(lxw_worksheet *worksheet,
                                      lxw_row_t row, lxw_col_t col,
                                      int value, lxw_format *format);

    lxw_error worksheet_write_blank(lxw_worksheet *worksheet,
                                    lxw_row_t row, lxw_col_t col,
                                    lxw_format *format);

    lxw_error worksheet_write_formula_num(lxw_worksheet *worksheet,
                                          lxw_row_t row,
                                          lxw_col_t col,
                                          const char *formula,
                                          lxw_format *format, double result);

    lxw_error worksheet_set_row(lxw_worksheet *worksheet,
                                lxw_row_t row, double height,
                                lxw_format *format);

    lxw_error worksheet_set_row_opt(lxw_worksheet *worksheet,
                                    lxw_row_t row,
                                    double height,
                                    lxw_format *format,
                                    lxw_row_col_options *options);

    lxw_error worksheet_set_column(lxw_worksheet *worksheet,
                                   lxw_col_t first_col,
                                   lxw_col_t last_col,
                                   double width, lxw_format *format);

    lxw_error worksheet_set_column_opt(lxw_worksheet *worksheet,
                                       lxw_col_t first_col,
                                       lxw_col_t last_col,
                                       double width,
                                       lxw_format *format,
                                       lxw_row_col_options *options);

    lxw_error worksheet_insert_image(lxw_worksheet *worksheet,
                                     lxw_row_t row, lxw_col_t col,
                                     const char *filename);

    lxw_error worksheet_insert_image_opt(lxw_worksheet *worksheet,
                                         lxw_row_t row, lxw_col_t col,
                                         const char *filename,
                                         lxw_image_options *options);

    lxw_error worksheet_insert_chart(lxw_worksheet *worksheet,
                                     lxw_row_t row, lxw_col_t col,
                                     lxw_chart *chart);

    lxw_error worksheet_insert_chart_opt(lxw_worksheet *worksheet,
                                         lxw_row_t row, lxw_col_t col,
                                         lxw_chart *chart,
                                         lxw_image_options *user_options);

    lxw_error worksheet_merge_range(lxw_worksheet *worksheet,
                                    lxw_row_t first_row,
                                    lxw_col_t first_col,
                                    lxw_row_t last_row,
                                    lxw_col_t last_col,
                                    const char *string,
                                    lxw_format *format);

    lxw_error worksheet_autofilter(lxw_worksheet *worksheet,
                                   lxw_row_t first_row,
                                   lxw_col_t first_col,
                                   lxw_row_t last_row,
                                   lxw_col_t last_col);

    void worksheet_activate(lxw_worksheet *worksheet);

    void worksheet_select(lxw_worksheet *worksheet);

    void worksheet_hide(lxw_worksheet *worksheet);

    void worksheet_set_first_sheet(lxw_worksheet *worksheet);

    void worksheet_freeze_panes(lxw_worksheet *worksheet,
                                lxw_row_t row, lxw_col_t col);

    void worksheet_split_panes(lxw_worksheet *worksheet,
                               double vertical, double horizontal);

    void worksheet_freeze_panes_opt(lxw_worksheet *worksheet,
                                    lxw_row_t first_row, lxw_col_t first_col,
                                    lxw_row_t top_row, lxw_col_t left_col,
                                    uint8_t type);

    void worksheet_split_panes_opt(lxw_worksheet *worksheet,
                                   double vertical, double horizontal,
                                   lxw_row_t top_row, lxw_col_t left_col);

    void worksheet_set_selection(lxw_worksheet *worksheet,
                                 lxw_row_t first_row, lxw_col_t first_col,
                                 lxw_row_t last_row, lxw_col_t last_col);

    void worksheet_set_landscape(lxw_worksheet *worksheet);

    void worksheet_set_portrait(lxw_worksheet *worksheet);

    void worksheet_set_page_view(lxw_worksheet *worksheet);

    void worksheet_set_paper(lxw_worksheet *worksheet, uint8_t paper_type);

    void worksheet_set_margins(lxw_worksheet *worksheet, double left,
                               double right, double top, double bottom);

    lxw_error worksheet_set_header(lxw_worksheet *worksheet,
                                   const char *string);

    lxw_error worksheet_set_footer(lxw_worksheet *worksheet,
                                   const char *string);

    lxw_error worksheet_set_header_opt(lxw_worksheet *worksheet,
                                       const char *string,
                                       lxw_header_footer_options *options);

    lxw_error worksheet_set_footer_opt(lxw_worksheet *worksheet,
                                       const char *string,
                                       lxw_header_footer_options *options);

    lxw_error worksheet_set_h_pagebreaks(lxw_worksheet *worksheet,
                                         lxw_row_t breaks[]);

    lxw_error worksheet_set_v_pagebreaks(lxw_worksheet *worksheet,
                                         lxw_col_t breaks[]);

    void worksheet_print_across(lxw_worksheet *worksheet);

    void worksheet_set_zoom(lxw_worksheet *worksheet, uint16_t scale);

    void worksheet_gridlines(lxw_worksheet *worksheet, uint8_t option);

    void worksheet_center_horizontally(lxw_worksheet *worksheet);

    void worksheet_center_vertically(lxw_worksheet *worksheet);

    void worksheet_print_row_col_headers(lxw_worksheet *worksheet);

    lxw_error worksheet_repeat_rows(lxw_worksheet *worksheet,
                                    lxw_row_t first_row,
                                    lxw_row_t last_row);

    lxw_error worksheet_repeat_columns(lxw_worksheet *worksheet,
                                       lxw_col_t first_col,
                                       lxw_col_t last_col);

    lxw_error worksheet_print_area(lxw_worksheet *worksheet,
                                   lxw_row_t first_row,
                                   lxw_col_t first_col,
                                   lxw_row_t last_row,
                                   lxw_col_t last_col);

    void worksheet_fit_to_pages(lxw_worksheet *worksheet, uint16_t width,
                                uint16_t height);

    void worksheet_set_start_page(lxw_worksheet *worksheet,
                                  uint16_t start_page);

    void worksheet_set_print_scale(lxw_worksheet *worksheet, uint16_t scale);

    void worksheet_right_to_left(lxw_worksheet *worksheet);

    void worksheet_hide_zero(lxw_worksheet *worksheet);

    void worksheet_set_tab_color(lxw_worksheet *worksheet, lxw_color_t color);

    void worksheet_protect(lxw_worksheet *worksheet, const char *password,
                           lxw_protection *options);

    void worksheet_set_default_row(lxw_worksheet *worksheet, double height,
                                   uint8_t hide_unused_rows);

    lxw_worksheet *lxw_worksheet_new(lxw_worksheet_init_data *init_data);

    void lxw_worksheet_free(lxw_worksheet *worksheet);

    void lxw_worksheet_assemble_xml_file(lxw_worksheet *worksheet);

    void lxw_worksheet_write_single_row(lxw_worksheet *worksheet);

    void lxw_worksheet_prepare_image(lxw_worksheet *worksheet,
                                     uint16_t image_ref_id,
                                     uint16_t drawing_id,
                                     lxw_image_options *image_data);

    void lxw_worksheet_prepare_chart(lxw_worksheet *worksheet,
                                     uint16_t chart_ref_id,
                                     uint16_t drawing_id,
                                     lxw_image_options *image_data);

    lxw_row *lxw_worksheet_find_row(lxw_worksheet *worksheet,
                                    lxw_row_t row_num);

    lxw_cell *lxw_worksheet_find_cell(lxw_row *row, lxw_col_t col_num);


cdef extern from "../libxlsxwriter/src/worksheet.c" nogil:
    lxw_error worksheet_write_number(lxw_worksheet *self,
                                     lxw_row_t row_num,
                                     lxw_col_t col_num,
                                     double value,
                                     lxw_format *format)

    lxw_error worksheet_write_string(lxw_worksheet *self,
                                     lxw_row_t row_num,
                                     lxw_col_t col_num,
                                     const char *string,
                                     lxw_format *format)

    lxw_error worksheet_write_formula_num(lxw_worksheet *self,
                                          lxw_row_t row_num,
                                          lxw_col_t col_num,
                                          const char *formula,
                                          lxw_format *format,
                                          double result)

    lxw_error worksheet_write_formula(lxw_worksheet *self,
                                      lxw_row_t row_num,
                                      lxw_col_t col_num,
                                      const char *formula,
                                      lxw_format *format)

    lxw_error worksheet_write_array_formula_num(lxw_worksheet *self,
                                                lxw_row_t first_row,
                                                lxw_col_t first_col,
                                                lxw_row_t last_row,
                                                lxw_col_t last_col,
                                                const char *formula,
                                                lxw_format *format,
                                                double result)

    lxw_error worksheet_write_array_formula(lxw_worksheet *self,
                                            lxw_row_t first_row,
                                            lxw_col_t first_col,
                                            lxw_row_t last_row,
                                            lxw_col_t last_col,
                                            const char *formula,
                                            lxw_format *format)
    lxw_error worksheet_write_blank(lxw_worksheet *self,
                                    lxw_row_t row_num,
                                    lxw_col_t col_num,
                                    lxw_format *format)

    lxw_error worksheet_write_boolean(lxw_worksheet *self,
                                      lxw_row_t row_num,
                                      lxw_col_t col_num,
                                      int value,
                                      lxw_format *format)

    lxw_error worksheet_write_datetime(lxw_worksheet *self,
                                       lxw_row_t row_num,
                                       lxw_col_t col_num,
                                       lxw_datetime *datetime,
                                       lxw_format *format)

    lxw_error worksheet_write_url_opt(lxw_worksheet *self,
                                      lxw_row_t row_num,
                                      lxw_col_t col_num,
                                      const char *url,
                                      lxw_format *format,
                                      const char *string,
                                      const char *tooltip)

    lxw_error worksheet_write_url(lxw_worksheet *self,
                                  lxw_row_t row_num,
                                  lxw_col_t col_num,
                                  const char *url,
                                  lxw_format *format)

    lxw_error worksheet_set_column_opt(lxw_worksheet *self,
                                       lxw_col_t firstcol,
                                       lxw_col_t lastcol,
                                       double width,
                                       lxw_format *format,
                                       lxw_row_col_options *user_options)

    lxw_error worksheet_set_column(lxw_worksheet *self,
                                   lxw_col_t firstcol,
                                   lxw_col_t lastcol,
                                   double width,
                                   lxw_format *format)

    lxw_error worksheet_set_row_opt(lxw_worksheet *self,
                                    lxw_row_t row_num,
                                    double height,
                                    lxw_format *format,
                                    lxw_row_col_options *user_options)

    lxw_error worksheet_set_row(lxw_worksheet *self,
                                lxw_row_t row_num,
                                double height,
                                lxw_format *format)

    lxw_error worksheet_merge_range(lxw_worksheet *self,
                                    lxw_row_t first_row,
                                    lxw_col_t first_col,
                                    lxw_row_t last_row,
                                    lxw_col_t last_col,
                                    const char *string,
                                    lxw_format *format)

    lxw_error worksheet_autofilter(lxw_worksheet *self,
                                   lxw_row_t first_row,
                                   lxw_col_t first_col, lxw_row_t last_row,
                                   lxw_col_t last_col)

    void worksheet_select(lxw_worksheet *self)

    void worksheet_activate(lxw_worksheet *self)

    void worksheet_set_first_sheet(lxw_worksheet *self)

    void worksheet_hide(lxw_worksheet *self)

    void worksheet_set_selection(lxw_worksheet *self,
                                 lxw_row_t first_row,
                                 lxw_col_t first_col,
                                 lxw_row_t last_row,
                                 lxw_col_t last_col)

    void worksheet_freeze_panes_opt(lxw_worksheet *self,
                                    lxw_row_t first_row,
                                    lxw_col_t first_col,
                                    lxw_row_t top_row,
                                    lxw_col_t left_col,
                                    uint8_t type)

    void worksheet_freeze_panes(lxw_worksheet *self,
                                lxw_row_t first_row,
                                lxw_col_t first_col)

    void worksheet_split_panes_opt(lxw_worksheet *self,
                                   double y_split,
                                   double x_split,
                                   lxw_row_t top_row,
                                   lxw_col_t left_col)

    void worksheet_split_panes(lxw_worksheet *self,
                               double y_split,
                               double x_split)

    void worksheet_set_portrait(lxw_worksheet *self)

    void worksheet_set_landscape(lxw_worksheet *self)

    void worksheet_set_page_view(lxw_worksheet *self)

    void worksheet_set_paper(lxw_worksheet *self, uint8_t paper_size)

    void worksheet_print_across(lxw_worksheet *self)

    void worksheet_set_margins(lxw_worksheet *self,
                               double left,
                               double right,
                               double top,
                               double bottom)

    lxw_error worksheet_set_header_opt(lxw_worksheet *self,
                                       const char *string,
                                       lxw_header_footer_options *options)

    lxw_error worksheet_set_footer_opt(lxw_worksheet *self,
                                       const char *string,
                                       lxw_header_footer_options *options)

    lxw_error worksheet_set_header(lxw_worksheet *self, const char *string)

    lxw_error worksheet_set_footer(lxw_worksheet *self, const char *string)

    void worksheet_gridlines(lxw_worksheet *self, uint8_t option)

    void worksheet_center_horizontally(lxw_worksheet *self)

    void worksheet_center_vertically(lxw_worksheet *self)

    void worksheet_print_row_col_headers(lxw_worksheet *self)

    lxw_error worksheet_repeat_rows(lxw_worksheet *self,
                                    lxw_row_t first_row,
                                    lxw_row_t last_row)

    lxw_error worksheet_repeat_columns(lxw_worksheet *self,
                                       lxw_col_t first_col,
                                       lxw_col_t last_col)

    lxw_error worksheet_print_area(lxw_worksheet *self,
                                   lxw_row_t first_row,
                                   lxw_col_t first_col,
                                   lxw_row_t last_row,
                                   lxw_col_t last_col)

    void worksheet_fit_to_pages(lxw_worksheet *self, uint16_t width, uint16_t height)

    void worksheet_set_start_page(lxw_worksheet *self, uint16_t start_page)

    void worksheet_set_print_scale(lxw_worksheet *self, uint16_t scale)

    lxw_error worksheet_set_h_pagebreaks(lxw_worksheet *self, lxw_row_t hbreaks[])

    lxw_error worksheet_set_v_pagebreaks(lxw_worksheet *self, lxw_col_t vbreaks[])

    void worksheet_set_zoom(lxw_worksheet *self, uint16_t scale)

    void worksheet_hide_zero(lxw_worksheet *self)

    void worksheet_right_to_left(lxw_worksheet *self)

    void worksheet_set_tab_color(lxw_worksheet *self, lxw_color_t color)

    void worksheet_protect(lxw_worksheet *self,
                           const char *password,
                           lxw_protection *options)

    void worksheet_set_default_row(lxw_worksheet *self,
                                   double height,
                                   uint8_t hide_unused_rows)

    lxw_error worksheet_insert_image_opt(lxw_worksheet *self,
                                         lxw_row_t row_num,
                                         lxw_col_t col_num,
                                         const char *filename,
                                         lxw_image_options *user_options)

    lxw_error worksheet_insert_image(lxw_worksheet *self,
                                     lxw_row_t row_num,
                                     lxw_col_t col_num,
                                     const char *filename)

    lxw_error worksheet_insert_chart_opt(lxw_worksheet *self,
                                         lxw_row_t row_num,
                                         lxw_col_t col_num,
                                         lxw_chart *chart,
                                         lxw_image_options *user_options)

    lxw_error worksheet_insert_chart(lxw_worksheet *self,
                                     lxw_row_t row_num,
                                     lxw_col_t col_num,
                                     lxw_chart *chart)


cdef extern from "../libxlsxwriter/include/xlsxwriter/utility.h" nogil:
    unsigned int lxw_name_to_row(const char *row_str);

    unsigned short lxw_name_to_col(const char *col_str);

    unsigned int lxw_name_to_row_2(const char *row_str);

    unsigned short lxw_name_to_col_2(const char *col_str);


cdef extern from "../libxlsxwriter/src/utility.c" nogil:
    ctypedef unsigned int lxw_row_t;

    ctypedef unsigned short lxw_col_t;

    lxw_row_t lxw_name_to_row(const char *row_str);

    lxw_col_t lxw_name_to_col(const char *col_str);

    unsigned int lxw_name_to_row_2(const char *row_str);

    unsigned short lxw_name_to_col_2(const char *col_str);


cdef extern from "../libxlsxwriter/include/xlsxwriter/shared_strings.h" nogil:
    ctypedef struct lxw_rb_generate_element:

        pass
    ctypedef struct sst_element:

        pass
    ctypedef struct lxw_sst:

        pass

    lxw_sst *lxw_sst_new;

    void lxw_sst_free(lxw_sst *sst);

    ctypedef struct sst_element:
        pass

    void lxw_sst_assemble_xml_file(lxw_sst *self);

    void _sst_xml_declaration(lxw_sst *self);


cdef extern from "../libxlsxwriter/src/shared_strings.c" nogil:
    lxw_sst *lxw_sst_new;

    void lxw_sst_free(lxw_sst *sst);

    void lxw_sst_assemble_xml_file(lxw_sst *self)

    ctypedef struct sst_element:
        pass


cdef extern from "../libxlsxwriter/include/xlsxwriter/hash_table.h" nogil:
    ctypedef struct lxw_hash_table:
        pass

    ctypedef struct lxw_hash_element:
        pass

    lxw_hash_element *lxw_hash_key_exists(lxw_hash_table *lxw_hash,
                                          void *key,
                                          size_t key_len);

    lxw_hash_element *lxw_insert_hash_element(lxw_hash_table *lxw_hash,
                                              void *key,
                                              void *value,
                                              size_t key_len);

    lxw_hash_table *lxw_hash_new(uint32_t num_buckets,
                                 uint8_t free_key,
                                 uint8_t free_value);

    void lxw_hash_free(lxw_hash_table *lxw_hash);


cdef extern from "../libxlsxwriter/src/hash_table.c" nogil:
    lxw_hash_element *lxw_hash_key_exists(lxw_hash_table *lxw_hash,
                                          void *key,
                                          size_t key_len);

    lxw_hash_element *lxw_insert_hash_element(lxw_hash_table *lxw_hash,
                                              void *key,
                                              void *value,
                                              size_t key_len);

    lxw_hash_table *lxw_hash_new(uint32_t num_buckets,
                                 uint8_t free_key,
                                 uint8_t free_value);

    void lxw_hash_free(lxw_hash_table *lxw_hash);


cdef extern from "../libxlsxwriter/include/xlsxwriter/format.h" nogil:
    ctypedef struct lxw_color_t:
        pass

    enum lxw_format_underlines:
        pass

    enum lxw_format_scripts:
        pass

    enum lxw_format_alignments:
        pass

    enum lxw_format_diagonal_types:
        pass

    enum lxw_defined_colors:
        pass

    enum lxw_format_patterns:
        pass

    enum lxw_format_borders:
        pass

    ctypedef struct lxw_format:
        pass

    ctypedef struct lxw_font:
        pass

    ctypedef struct lxw_border:
        pass

    ctypedef struct lxw_fill:
        pass

    lxw_format *lxw_format_new();

    void lxw_format_free(lxw_format *format);

    int lxw_format_get_xf_index(lxw_format *format);

    lxw_font *lxw_format_get_font_key(lxw_format *format);

    lxw_border *lxw_format_get_border_key(lxw_format *format);

    lxw_fill *lxw_format_get_fill_key(lxw_format *format);

    lxw_color_t lxw_format_check_color(lxw_color_t color);

    void format_set_font_name(lxw_format *format, const char *font_name);

    void format_set_font_size(lxw_format *format, uint16_t size);

    void format_set_font_color(lxw_format *format, lxw_color_t color);

    void format_set_bold(lxw_format *format);

    void format_set_italic(lxw_format *format);

    void format_set_underline(lxw_format *format, uint8_t style);

    void format_set_font_strikeout(lxw_format *format);

    void format_set_font_script(lxw_format *format, uint8_t style);

    void format_set_num_format(lxw_format *format, const char *num_format);

    void format_set_num_format_index(lxw_format *format, uint8_t index);

    void format_set_unlocked(lxw_format *format);

    void format_set_hidden(lxw_format *format);

    void format_set_align(lxw_format *format, uint8_t alignment);

    void format_set_text_wrap(lxw_format *format);

    void format_set_rotation(lxw_format *format, int16_t angle);

    void format_set_indent(lxw_format *format, uint8_t level);

    void format_set_shrink(lxw_format *format);

    void format_set_pattern(lxw_format *format, uint8_t index);

    void format_set_bg_color(lxw_format *format, lxw_color_t color);

    void format_set_fg_color(lxw_format *format, lxw_color_t color);

    void format_set_border(lxw_format *format, uint8_t style);

    void format_set_bottom(lxw_format *format, uint8_t style);

    void format_set_top(lxw_format *format, uint8_t style);

    void format_set_left(lxw_format *format, uint8_t style);

    void format_set_right(lxw_format *format, uint8_t style);

    void format_set_border_color(lxw_format *format, lxw_color_t color);

    void format_set_bottom_color(lxw_format *format, lxw_color_t color);

    void format_set_top_color(lxw_format *format, lxw_color_t color);

    void format_set_left_color(lxw_format *format, lxw_color_t color);

    void format_set_right_color(lxw_format *format, lxw_color_t color);

    void format_set_diag_type(lxw_format *format, uint8_t value);

    void format_set_diag_color(lxw_format *format, lxw_color_t color);

    void format_set_diag_border(lxw_format *format, uint8_t value);

    void format_set_font_outline(lxw_format *format);

    void format_set_font_shadow(lxw_format *format);

    void format_set_font_family(lxw_format *format, uint8_t value);

    void format_set_font_charset(lxw_format *format, uint8_t value);

    void format_set_font_scheme(lxw_format *format, const char *font_scheme);

    void format_set_font_condense(lxw_format *format);

    void format_set_font_extend(lxw_format *format);

    void format_set_reading_order(lxw_format *format, uint8_t value);

    void format_set_theme(lxw_format *format, uint8_t value);


cdef extern from "../libxlsxwriter/src/format.c" nogil:
    lxw_format *lxw_format_new;

    void lxw_format_free(lxw_format *format);

    lxw_color_t lxw_format_check_color(lxw_color_t color);

    lxw_font *lxw_format_get_font_key(lxw_format *self);

    lxw_border *lxw_format_get_border_key(lxw_format *self);

    lxw_fill *lxw_format_get_fill_key(lxw_format *self);

    int lxw_format_get_xf_index(lxw_format *self);

    void format_set_font_name(lxw_format *self, const char *font_name);

    void format_set_font_size(lxw_format *self, uint16_t size);

    void format_set_font_color(lxw_format *self, lxw_color_t color);

    void format_set_bold(lxw_format *self);

    void format_set_italic(lxw_format *self);

    void format_set_underline(lxw_format *self, uint8_t style);

    void format_set_font_strikeout(lxw_format *self);

    void format_set_font_script(lxw_format *self, uint8_t style);

    void format_set_font_outline(lxw_format *self);

    void format_set_font_shadow(lxw_format *self);

    void format_set_num_format(lxw_format *self, const char *num_format);

    void format_set_unlocked(lxw_format *self);

    void format_set_hidden(lxw_format *self);

    void format_set_align(lxw_format *self, uint8_t value);

    void format_set_text_wrap(lxw_format *self);

    void format_set_rotation(lxw_format *self, int16_t angle);

    void format_set_indent(lxw_format *self, uint8_t value);

    void format_set_shrink(lxw_format *self);

    void format_set_text_justlast(lxw_format *self);

    void format_set_pattern(lxw_format *self, uint8_t value);

    void format_set_bg_color(lxw_format *self, lxw_color_t color);

    void format_set_fg_color(lxw_format *self, lxw_color_t color);

    void format_set_border(lxw_format *self, uint8_t style);

    void format_set_border_color(lxw_format *self, lxw_color_t color);

    void format_set_bottom(lxw_format *self, uint8_t style);

    void format_set_bottom_color(lxw_format *self, lxw_color_t color);

    void format_set_left(lxw_format *self, uint8_t style);

    void format_set_left_color(lxw_format *self, lxw_color_t color);

    void format_set_right(lxw_format *self, uint8_t style);

    void format_set_right_color(lxw_format *self, lxw_color_t color);

    void format_set_top(lxw_format *self, uint8_t style);

    void format_set_top_color(lxw_format *self, lxw_color_t color);

    void format_set_diag_type(lxw_format *self, uint8_t type);

    void format_set_diag_color(lxw_format *self, lxw_color_t color);

    void format_set_diag_border(lxw_format *self, uint8_t style);

    void format_set_num_format_index(lxw_format *self, uint8_t value);

    void format_set_valign(lxw_format *self, uint8_t value);

    void format_set_reading_order(lxw_format *self, uint8_t value);

    void format_set_font_family(lxw_format *self, uint8_t value);

    void format_set_font_charset(lxw_format *self, uint8_t value);

    void format_set_font_scheme(lxw_format *self, const char *font_scheme);

    void format_set_font_condense(lxw_format *self);

    void format_set_font_extend(lxw_format *self);

    void format_set_theme(lxw_format *self, uint8_t value);


cdef extern from "../libxlsxwriter/include/xlsxwriter/chart.h" nogil:
    ctypedef struct lxw_chart:
        pass

    lxw_chart *lxw_chart_new(uint8_t type);

    void lxw_chart_free(lxw_chart *chart);


cdef extern from "../libxlsxwriter/src/chart.c" nogil:
    void lxw_chart_free(lxw_chart *chart)

    lxw_chart * lxw_chart_new(uint8_t type)

