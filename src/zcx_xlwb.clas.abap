CLASS zcx_xlwb DEFINITION
  PUBLIC
  INHERITING FROM cx_static_check
  FINAL
  CREATE PUBLIC.

  PUBLIC SECTION.
    DATA mv_text TYPE string READ-ONLY.

    METHODS constructor
      IMPORTING
        iv_text  TYPE string OPTIONAL
        previous LIKE previous OPTIONAL.

    METHODS get_text REDEFINITION.
ENDCLASS.



CLASS ZCX_XLWB IMPLEMENTATION.


  METHOD constructor ##ADT_SUPPRESS_GENERATION.
    super->constructor( previous = previous ).
    mv_text = iv_text.
  ENDMETHOD.


  METHOD get_text.
    IF mv_text IS NOT INITIAL.
      result = mv_text.
    ELSE.
      result = super->get_text( ).
    ENDIF.
  ENDMETHOD.
ENDCLASS.
