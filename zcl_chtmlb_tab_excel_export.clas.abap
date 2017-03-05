class ZCL_CHTMLB_TAB_EXCEL_EXPORT definition
  public
  create public .

public section.
*"* public components of class ZCL_CHTMLB_TAB_EXCEL_EXPORT
*"* do not include other source files here!!!

  interfaces IF_HTTP_EXTENSION .
protected section.
*"* protected components of class ZCL_CHTMLB_TAB_EXCEL_EXPORT
*"* do not include other source files here!!!

  data CONFIG_TABLE type BSP_DLC_COLUMN_DESCR_TAB .
  data IO_EXCEL type ref to ZCL_EXCEL .
  data IO_TV_CONTEXT_NODE type ref to CL_BSP_WD_CONTEXT_NODE_TV .
  data IO_VIEW_CONTROLLER type ref to CL_BSP_WD_VIEW_CONTROLLER .
  data IO_VIEW_DESCRIPTOR type ref to IF_BSP_DLC_VIEW_DESCRIPTOR .
  data IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET .
  data IS_TREE type ABAP_BOOL .
  data TAB_SIZE type I .
  data IO_STYLE_1 type ref to ZCL_EXCEL_STYLE .
  data IO_STYLE_2 type ref to ZCL_EXCEL_STYLE .
  data IO_STYLE_3 type ref to ZCL_EXCEL_STYLE .

  methods FILL_TABLE_CONTENT .
  methods FILL_TABLE_HEADER .
  methods GET_CELL_STYLE
    importing
      !IV_COLUMN type I optional
      !IV_ROW type I optional
    returning
      value(RV_STYLE) type ZEXCEL_CELL_STYLE .
  methods SET_WORKSHEET_OPTIONS .
private section.
*"* private components of class ZCL_CHTMLB_TAB_EXCEL_EXPORT
*"* do not include other source files here!!!

  data IO_MODEL_TREE type ref to CL_BSP_WD_CONTEXT_NODE_TREE .

  methods CALC_TREE_INDENT
    importing
      !IR_TREE_MODEL type ref to CL_BSP_WD_CONTEXT_NODE_TREE
      !IV_INDEX type I
    returning
      value(RV_RETURN) type STRING .
ENDCLASS.



CLASS ZCL_CHTMLB_TAB_EXCEL_EXPORT IMPLEMENTATION.


method CALC_TREE_INDENT.
  DATA: node     TYPE crmt_thtmlb_treetable_node,
        original TYPE crmt_thtmlb_treetable_node,
        parent   TYPE crmt_thtmlb_treetable_node,
        indent   TYPE string.

  " get node key
  READ TABLE ir_tree_model->node_tab INDEX iv_index INTO node.
  original = node.

  WHILE node-parent_key IS NOT INITIAL.
    " walk up in the tree to get the root node
    CONCATENATE indent '--' INTO indent RESPECTING BLANKS. "EC NOTEXT
    READ TABLE ir_tree_model->node_tab INTO parent
               WITH KEY node_key = node-parent_key.
    node = parent.
  ENDWHILE.

  IF original-is_leaf IS INITIAL.
    IF original-is_expanded IS INITIAL.
      CONCATENATE indent '>' INTO indent. "EC NOTEXT
    ELSE.
      CONCATENATE indent 'v' INTO indent. "EC NOTEXT
    ENDIF.
  ELSE.
    CONCATENATE indent 'o' INTO indent. "EC NOTEXT
  ENDIF.

  rv_return = indent.
endmethod.


METHOD fill_table_content.
  DATA: dref TYPE REF TO data.
  FIELD-SYMBOLS: <line> TYPE any.

  FIELD-SYMBOLS: <config_line> LIKE LINE OF me->config_table.

  DATA: column                  TYPE zexcel_cell_column VALUE 1,
        row                     TYPE zexcel_cell_row,
        column_str              TYPE zexcel_cell_column_alpha,
        abap_type               TYPE abap_typekind.

  DATA: field_name    TYPE string,
        path          TYPE string,
        lo_metadata   TYPE REF TO if_bsp_metadata,
        col_type      TYPE sychar01.

  DATA: ex TYPE REF TO cx_root.
  " get structure of the table
  dref = io_tv_context_node->get_table_line_sample_ext( ).
  " Needed to provide abap2xlsx with correct data types
  ASSIGN dref->* TO <line>.

  LOOP AT me->config_table ASSIGNING <config_line>.
    IF <config_line>-mandatory IS INITIAL.
      field_name = <config_line>-name.
      CONCATENATE `TABLE[1].` field_name INTO path.
      TRY.
          lo_metadata = io_tv_context_node->get_m_t_table( attribute_path = path
                                                              component      = field_name ).
          col_type = lo_metadata->get_abap_type( ).
        CATCH: cx_root.
          col_type = '-'.
      ENDTRY.
      " (mis)use mandatory for the coltype
      IF col_type = 'C'.
        col_type = cl_abap_typedescr=>typekind_string.
      ENDIF.
      <config_line>-mandatory = col_type.
    ENDIF.
  ENDLOOP.

  " get table size
  DATA: tree_indent   TYPE string,
        textlen       TYPE i.

  DATA: index          TYPE i,
        field_index    TYPE i,
        length         TYPE i,
        index_s        TYPE string,
        cell_value     TYPE string.

  DATA: lo_valuehelp_descr     TYPE REF TO if_bsp_wd_valuehelp_descriptor,
        lo_picklist_descr      TYPE REF TO if_bsp_wd_valuehelp_pldescr,
        lt_picklist_values     TYPE bsp_wd_dropdown_table.

  DATA: pl_binding_string  TYPE string,
        pl_model_name      TYPE string,
        pl_mod             TYPE REF TO if_bsp_model_binding,
        pl_object          TYPE REF TO cl_bsp_wd_context_node_ddlb,
        lt_pl_entries         TYPE REF TO bsp_wd_dropdown_table.

  FIELD-SYMBOLS: <ls_ddlb_line> TYPE bsp_wd_dropdown_line,
                 <lt_entries>   TYPE bsp_wd_dropdown_table,
                 <cell_value>   TYPE any.

  DATA: field_type TYPE string,
        tooltip    TYPE string.
  " Read user settings for Decimal notation
  DATA: dec_point_format TYPE xudcpfm.
  SELECT SINGLE dcpfm "Decimal format
    FROM usr01
    INTO (dec_point_format)
    WHERE bname = sy-uname.

  row = 2.
  DO tab_size TIMES.
    column = 1.
    index = sy-index.

    " Get style
    DATA(lv_style) = get_cell_style( iv_row = row iv_column = column ).

    LOOP AT me->config_table ASSIGNING <config_line>.
      field_index = sy-tabix.
      abap_type  = <config_line>-mandatory.
      " check for indent if we export a tree
      IF field_index = 1.
        IF is_tree = abap_true.
          tree_indent = calc_tree_indent( ir_tree_model = io_model_tree
                                          iv_index      = index ).
        ENDIF.
      ENDIF.
      " get data for each cell
      field_name = <config_line>-name.
      index_s = index.
      CONCATENATE `TABLE[` index_s `].` field_name INTO path.
      TRY.
          cell_value = io_tv_context_node->get_t_table( attribute_path = path
                                                           index          = index
                                                           component      = field_name ).
        CATCH cx_root.
          CONCATENATE field_name ` not bound` INTO cell_value. "#EC NOTEXT
      ENDTRY.
      " check for special cell types
      " look for a picklist descriptor for this
      CLEAR lo_picklist_descr.
      TRY.
          lo_valuehelp_descr = io_tv_context_node->get_v_t_table( iv_index  = index
                                                                     component = field_name ).
          IF lo_valuehelp_descr IS BOUND.
            lo_picklist_descr ?= lo_valuehelp_descr.
          ENDIF.
        CATCH cx_root.
      ENDTRY.
      " if a picklist descriptor exists, use the description instead of the value
      IF lo_picklist_descr IS BOUND.
        CASE lo_picklist_descr->source_type.
          WHEN if_bsp_wd_valuehelp_pldescr=>source_type_binding.
            " picklist values through binding string
            pl_binding_string = lo_picklist_descr->get_binding_string( ).
            cl_bsp_model=>if_bsp_model_util~split_binding_expression(
              EXPORTING
                binding_expression = pl_binding_string
              IMPORTING
                model_name         = pl_model_name ).
            " get model reference
            TRY.
                pl_mod   = io_view_controller->get_model( pl_model_name ).
              CATCH: cx_bsp_inv_attr_name.
                EXIT.
            ENDTRY.
            pl_object ?= pl_mod.
            lt_pl_entries = pl_object->get_t_values( ).
            ASSIGN lt_pl_entries->* TO <lt_entries>.
            lt_picklist_values = <lt_entries>.

          WHEN if_bsp_wd_valuehelp_pldescr=>source_type_table.
            lt_picklist_values = lo_picklist_descr->get_selection_table( ).
        ENDCASE.
        READ TABLE lt_picklist_values WITH KEY key = cell_value
          ASSIGNING <ls_ddlb_line>.                         "#EC WARNOK
        IF sy-subrc IS INITIAL.
          cell_value = <ls_ddlb_line>-value.
        ELSE.
          IF sy-subrc = 4 AND cell_value IS INITIAL.
            cell_value = ''.
          ENDIF.
        ENDIF.
      ENDIF.

      " check for image fields
      field_type = io_tv_context_node->get_p_t_table( iv_index = index
                                                     component = field_name
                                                     iv_property = if_bsp_wd_model_setter_getter=>fp_fieldtype ).
      IF field_type = if_bsp_dlc_view_descriptor=>field_type_image.
        " get tooltip
        tooltip = io_tv_context_node->get_p_t_table( iv_index  = index
                                                    component = field_name
                                                    iv_property = if_bsp_wd_model_setter_getter=>fp_tooltip ).
        cell_value = tooltip.
      ENDIF.
      " remove line breaks
      REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>cr_lf IN cell_value WITH ' '.
      IF <config_line>-mandatory CA 'IPF'.
        " if number, convert to common format
        CASE dec_point_format.
          WHEN 'X'. " 1,234,567.89
            REPLACE ALL OCCURRENCES OF ',' IN cell_value WITH space.
            CONDENSE cell_value.
          WHEN OTHERS. " ' ' 1.234.567,89 or 'Y' 1 234 567,89
            REPLACE ALL OCCURRENCES OF '.' IN cell_value WITH space.
            CONDENSE cell_value.
            REPLACE ',' IN cell_value WITH '.'.
        ENDCASE.
        " Put - in front
        textlen = strlen( cell_value ).
        textlen = textlen - 1.
        IF textlen > 0.
          IF cell_value+textlen = '-'.
            CONCATENATE '-' cell_value(textlen) INTO cell_value.
          ENDIF.
        ENDIF.
      ELSEIF <config_line>-mandatory = cl_abap_typedescr=>typekind_date
         AND NOT cell_value IS INITIAL.
        CALL FUNCTION 'CONVERT_DATE_TO_INTERNAL'
          EXPORTING
            date_external = cell_value
          IMPORTING
            date_internal = cell_value.
      ELSEIF <config_line>-mandatory = cl_abap_typedescr=>typekind_time
         AND NOT cell_value IS INITIAL.
        CALL FUNCTION 'CONVERT_TIME_INPUT'
          EXPORTING
            input  = cell_value
          IMPORTING
            output = cell_value.
      ENDIF.

      " If we have a tree we have an extra column with the current level
      IF    field_index = 1
        AND is_tree     = abap_true.
        column_str = zcl_excel_common=>convert_column2alpha( column ).
        io_worksheet->set_cell(
          EXPORTING
            ip_column    =   column_str  " Cell Column
            ip_row       =   row         " Cell Row
            ip_value     =   tree_indent " Cell Value
            ip_data_type =   's'         " Excel type
            ip_style     = lv_style
        ).
        column = column + 1.
      ENDIF.

      column_str = zcl_excel_common=>convert_column2alpha( column ).

      ASSIGN COMPONENT <config_line>-name OF STRUCTURE <line> TO <cell_value>.
      DESCRIBE FIELD <cell_value> OUTPUT-LENGTH length.

      " When the cell_value is longer than the structure fild,
      " then it was translated from the key to the description
      IF length < strlen( cell_value )
      OR cell_value CN '0123456789., '. " Quick fix for issue #2

        io_worksheet->set_cell(
          EXPORTING
            ip_column    = column_str  " Cell Column
            ip_row       = row         " Cell Row
            ip_value     = cell_value  " Cell Value
            ip_data_type = 's'         " Excel type
            ip_style     = lv_style
        ).
      ELSE.
        TRY.
            <cell_value> = cell_value.
            io_worksheet->set_cell(
              EXPORTING
                ip_column    = column_str   " Cell Column
                ip_row       = row          " Cell Row
                ip_value     = <cell_value> " Cell Value
                ip_abap_type = abap_type    " ABAP cell data type
                ip_style     = lv_style
            ).
          CATCH cx_sy_conversion_no_number INTO ex.
            io_worksheet->set_cell(
              EXPORTING
                ip_column    = column_str  " Cell Column
                ip_row       = row         " Cell Row
                ip_value     = cell_value  " Cell Value
                ip_data_type = 's'         " Excel type
                ip_style     = lv_style
            ).
        ENDTRY.
      ENDIF.
      column = column + 1.
    ENDLOOP.
    row = row + 1.
  ENDDO.

ENDMETHOD.


method FILL_TABLE_HEADER.

  DATA: comp                TYPE cl_chtmlb_xml_provider=>chtmlb_tv_comp_descr,
        title               TYPE bsp_dlc_element_id,
        tmp_string          TYPE string,                    "#EC NEEDED
        column_dimension    TYPE REF TO zcl_excel_worksheet_columndime,
        column_str          TYPE zexcel_cell_column_alpha,
        column              TYPE zexcel_cell_column VALUE 1,
        lv_style_header     TYPE zexcel_cell_style.

  FIELD-SYMBOLS: <config_line> LIKE LINE OF me->config_table.
  " Creates active sheet
  CREATE OBJECT io_excel.

  " Get active sheet
  io_worksheet = io_excel->get_active_worksheet( ).
  io_worksheet->set_title( ip_title = 'CRM Table Export'(001) ).

  " Get header style
  lv_style_header = get_cell_style( iv_row = 1 ).

  " Is the table a tree?
  TRY.
      io_model_tree ?= io_tv_context_node.
      " tree
      is_tree = abap_true.
      tab_size = lines( io_model_tree->node_tab ).
      column_str = zcl_excel_common=>convert_column2alpha( column ).
      io_worksheet->set_cell(
        EXPORTING
          ip_column    =   column_str  " Cell Column
          ip_row       =   1  " Cell Row
          ip_value     =   'Level'(002)  " Cell Value
      ).
      column = column + 1.
    CATCH: cx_root.                                      "#EC CATCH_ALL
      " table
      is_tree = abap_false.
      tab_size = io_tv_context_node->get_table_size( ).
  ENDTRY.

  " Fill Table Header
  LOOP AT me->config_table ASSIGNING <config_line>.
    CLEAR comp.
    " read column name
    IF <config_line>-title(2) <> '//'.
      IF <config_line>-title IS NOT INITIAL.
        comp-column_title = <config_line>-title.
      ELSE.
        comp-column_title = <config_line>-name.
      ENDIF.
    ELSE.
      IF io_view_descriptor IS BOUND.
        " get field label
        title = <config_line>-title.
        comp-column_title = io_view_descriptor->get_field_label(
                                                      iv_element_id = title ).
      ENDIF.
    ENDIF.
    " remove 'STRUCT'
    WHILE <config_line>-name CA '.'.
      SPLIT <config_line>-name AT '.' INTO tmp_string <config_line>-name.
    ENDWHILE.
    comp-name = <config_line>-name.
    IF comp-column_title IS INITIAL.
      comp-column_title = <config_line>-name.
    ENDIF.
    column_str = zcl_excel_common=>convert_column2alpha( column ).
    io_worksheet->set_cell(
      EXPORTING
        ip_column    = column_str  " Cell Column
        ip_row       = 1  " Cell Row
        ip_value     = comp-column_title  " Cell Value
        ip_style     = lv_style_header
    ).
    " Autosize Columns
    column_dimension = io_worksheet->get_column_dimension( ip_column = column_str ).
    column_dimension->set_auto_size( abap_true ).

    column = column + 1.
  ENDLOOP.
endmethod.


METHOD get_cell_style.

  DATA lr_border_dark          TYPE REF TO zcl_excel_style_border.
  DATA lr_border_blue          TYPE REF TO zcl_excel_style_border.
  DATA lv_mod                  TYPE i.

  IF iv_row = 1.
    FREE: io_style_1, io_style_2, io_style_3.

    " Header line
    io_style_1                       = io_excel->add_new_style( ).
    io_style_1->fill->filltype       = zcl_excel_style_fill=>c_fill_solid.
    io_style_1->fill->fgcolor-theme  = zcl_excel_style_color=>c_theme_accent1.
    io_style_1->font->bold           = abap_true.
    io_style_1->font->color-rgb      = zcl_excel_style_color=>c_white.

    CREATE OBJECT lr_border_dark.
    lr_border_dark->border_color-rgb = zcl_excel_style_color=>c_black.
    lr_border_dark->border_style     = zcl_excel_style_border=>c_border_thin.
    io_style_1->borders->allborders  = lr_border_dark.

    rv_style = io_style_1->get_guid( ).

  ELSE.
    lv_mod = iv_row MOD 2.
    IF lv_mod = 0.
      " Table line: Pair
      IF io_style_2 IS INITIAL.
        io_style_2                         = io_excel->add_new_style( ).
        io_style_2->fill->filltype         = zcl_excel_style_fill=>c_fill_solid.
        io_style_2->fill->fgcolor-theme    = zcl_excel_style_color=>c_theme_accent1.
        io_style_2->fill->fgcolor-tint     = '0.79998168889431442'.
        io_style_2->fill->bgcolor-indexed  = zcl_excel_style_color=>c_indexed_sys_foreground.

        CREATE OBJECT lr_border_blue.
        lr_border_blue->border_color-theme = zcl_excel_style_color=>c_theme_accent1.
        lr_border_blue->border_style       = zcl_excel_style_border=>c_border_thin.
        io_style_2->borders->allborders    = lr_border_blue.
      ENDIF.
      rv_style = io_style_2->get_guid( ).

    ELSE.
      " Table line: Impair
      IF io_style_3 IS INITIAL.
        io_style_3                         = io_excel->add_new_style( ).

        CREATE OBJECT lr_border_blue.
        lr_border_blue->border_color-theme = zcl_excel_style_color=>c_theme_accent1.
        lr_border_blue->border_style       = zcl_excel_style_border=>c_border_thin.
        io_style_3->borders->allborders    = lr_border_blue.
      ENDIF.
      rv_style = io_style_3->get_guid( ).
    ENDIF.
  ENDIF.

ENDMETHOD.


method IF_HTTP_EXTENSION~HANDLE_REQUEST.
  DATA: request_method         TYPE string,
        instance_id            TYPE string.

  DATA: lo_excel_writer         TYPE REF TO zif_excel_writer,
        file                    TYPE xstring.

  " Get HTTP request method
  request_method = server->request->if_http_entity~get_header_field( `~request_method` ).
  CASE request_method.
      " Handle HTTP GET or POST method
    WHEN 'GET' OR
         'POST'.

      server->request->get_form_data(
        EXPORTING
          name = 'iId'
        CHANGING
          data = instance_id ).

      " get model for table
      io_tv_context_node ?= cl_chtmlb_config_tab_excel_exp=>model_get( instance_id ).
      CHECK io_tv_context_node IS BOUND.

      " get view controller for retreiving configuration
      io_view_controller ?= cl_chtmlb_config_tab_excel_exp=>controller_get( instance_id ).
      CHECK io_view_controller IS BOUND.

      " get view descriptor
      io_view_descriptor ?= io_view_controller->configuration_descr->get_property_descriptor( ).

      " get table configuration
      me->config_table = io_tv_context_node->get_table_components( ).
      " Return only displayed columns
      DELETE me->config_table WHERE hidden = abap_true.
      " Do not output One-Click-Action Column
      DELETE me->config_table WHERE name = 'THTMLB_OCA'.

      me->fill_table_header( ).
      me->fill_table_content( ).
      me->set_worksheet_options( ).

      CREATE OBJECT lo_excel_writer TYPE zcl_excel_writer_2007.
      file = lo_excel_writer->write_file( io_excel ).

      server->response->set_header_field( name  = 'Content-Type' "#EC NOTEXT
                                          value = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ). "#EC NOTEXT

      server->response->set_header_field( name  = 'expires' "#EC NOTEXT
                                          value = '0' ).    "#EC NOTEXT
      DATA: filename TYPE string.
      CONCATENATE 'attachment; filename="CRM-Download-' sy-datum '-' sy-uzeit '.xlsx' INTO filename. "#EC NOTEXT

      server->response->set_header_field( name  = 'Content-Disposition' "#EC NOTEXT
                                          value = filename ).


      server->response->set_data( data = file ).

    WHEN OTHERS.
  ENDCASE.
endmethod.


METHOD set_worksheet_options.

  DATA lr_autofilter           TYPE REF TO zcl_excel_autofilter.
  DATA ls_area                 TYPE zexcel_s_autofilter_area.

  TRY.
      io_worksheet->freeze_panes(
        EXPORTING
          ip_num_rows    = 1 ).

      ls_area-row_start = 1.
      ls_area-col_start = 1.
      ls_area-row_end   = io_worksheet->get_highest_row( ).
      ls_area-col_end   = io_worksheet->get_highest_column( ).

      lr_autofilter = io_excel->add_new_autofilter( io_sheet = io_worksheet ) .
      lr_autofilter->set_filter_area( is_area = ls_area ).

    CATCH zcx_excel.
      RETURN.
  ENDTRY.

ENDMETHOD.
ENDCLASS.
