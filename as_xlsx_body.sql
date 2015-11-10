CREATE OR REPLACE PACKAGE BODY as_xlsx IS
  --
  c_LOCAL_FILE_HEADER        CONSTANT RAW(4) := hextoraw('504B0304'); -- Local file header signature
  c_END_OF_CENTRAL_DIRECTORY CONSTANT RAW(4) := hextoraw('504B0506'); -- End of central directory signature
  --
  TYPE tp_XF_fmt IS RECORD(
     numFmtId  PLS_INTEGER
    ,fontId    PLS_INTEGER
    ,fillId    PLS_INTEGER
    ,borderId  PLS_INTEGER
    ,alignment tp_alignment);
  TYPE tp_col_fmts IS TABLE OF tp_XF_fmt INDEX BY PLS_INTEGER;
  TYPE tp_row_fmts IS TABLE OF tp_XF_fmt INDEX BY PLS_INTEGER;
  TYPE tp_widths IS TABLE OF NUMBER INDEX BY PLS_INTEGER;
  TYPE tp_cell IS RECORD(
     VALUE NUMBER
    ,style VARCHAR2(50));
  TYPE tp_cells IS TABLE OF tp_cell INDEX BY PLS_INTEGER;
  TYPE tp_rows IS TABLE OF tp_cells INDEX BY PLS_INTEGER;
  TYPE tp_autofilter IS RECORD(
     column_start PLS_INTEGER
    ,column_end   PLS_INTEGER
    ,row_start    PLS_INTEGER
    ,row_end      PLS_INTEGER);
  TYPE tp_autofilters IS TABLE OF tp_autofilter INDEX BY PLS_INTEGER;
  TYPE tp_hyperlink IS RECORD(
     cell VARCHAR2(10)
    ,url  VARCHAR2(1000));
  TYPE tp_hyperlinks IS TABLE OF tp_hyperlink INDEX BY PLS_INTEGER;
  SUBTYPE tp_author IS VARCHAR2(32767 CHAR);
  TYPE tp_authors IS TABLE OF PLS_INTEGER INDEX BY tp_author;
  authors tp_authors;
  TYPE tp_comment IS RECORD(
     text   VARCHAR2(32767 CHAR)
    ,author tp_author
    ,ROW    PLS_INTEGER
    ,column PLS_INTEGER
    ,width  PLS_INTEGER
    ,height PLS_INTEGER);
  TYPE tp_comments IS TABLE OF tp_comment INDEX BY PLS_INTEGER;
  TYPE tp_mergecells IS TABLE OF VARCHAR2(21) INDEX BY PLS_INTEGER;
  TYPE tp_validation IS RECORD(
     TYPE             VARCHAR2(10)
    ,errorstyle       VARCHAR2(32)
    ,showinputmessage BOOLEAN
    ,prompt           VARCHAR2(32767 CHAR)
    ,title            VARCHAR2(32767 CHAR)
    ,error_title      VARCHAR2(32767 CHAR)
    ,error_txt        VARCHAR2(32767 CHAR)
    ,showerrormessage BOOLEAN
    ,formula1         VARCHAR2(32767 CHAR)
    ,formula2         VARCHAR2(32767 CHAR)
    ,allowBlank       BOOLEAN
    ,sqref            VARCHAR2(32767 CHAR));
  TYPE tp_validations IS TABLE OF tp_validation INDEX BY PLS_INTEGER;
  TYPE tp_sheet IS RECORD(
     rows        tp_rows
    ,widths      tp_widths
    ,NAME        VARCHAR2(100)
    ,freeze_rows PLS_INTEGER
    ,freeze_cols PLS_INTEGER
    ,autofilters tp_autofilters
    ,hyperlinks  tp_hyperlinks
    ,col_fmts    tp_col_fmts
    ,row_fmts    tp_row_fmts
    ,comments    tp_comments
    ,mergecells  tp_mergecells
    ,validations tp_validations);
  TYPE tp_sheets IS TABLE OF tp_sheet INDEX BY PLS_INTEGER;
  TYPE tp_numFmt IS RECORD(
     numFmtId   PLS_INTEGER
    ,formatCode VARCHAR2(100));
  TYPE tp_numFmts IS TABLE OF tp_numFmt INDEX BY PLS_INTEGER;
  TYPE tp_fill IS RECORD(
     patternType VARCHAR2(30)
    ,fgRGB       VARCHAR2(8));
  TYPE tp_fills IS TABLE OF tp_fill INDEX BY PLS_INTEGER;
  TYPE tp_cellXfs IS TABLE OF tp_xf_fmt INDEX BY PLS_INTEGER;
  TYPE tp_font IS RECORD(
     NAME      VARCHAR2(100)
    ,family    PLS_INTEGER
    ,fontsize  NUMBER
    ,theme     PLS_INTEGER
    ,RGB       VARCHAR2(8)
    ,underline BOOLEAN
    ,italic    BOOLEAN
    ,bold      BOOLEAN);
  TYPE tp_fonts IS TABLE OF tp_font INDEX BY PLS_INTEGER;
  TYPE tp_border IS RECORD(
     top    VARCHAR2(17)
    ,bottom VARCHAR2(17)
    ,LEFT   VARCHAR2(17)
    ,RIGHT  VARCHAR2(17));
  TYPE tp_borders IS TABLE OF tp_border INDEX BY PLS_INTEGER;
  TYPE tp_numFmtIndexes IS TABLE OF PLS_INTEGER INDEX BY PLS_INTEGER;
  TYPE tp_strings IS TABLE OF PLS_INTEGER INDEX BY VARCHAR2(32767 CHAR);
  TYPE tp_str_ind IS TABLE OF VARCHAR2(32767 CHAR) INDEX BY PLS_INTEGER;
  TYPE tp_defined_name IS RECORD(
     NAME  VARCHAR2(32767 CHAR)
    ,REF   VARCHAR2(32767 CHAR)
    ,sheet PLS_INTEGER);
  TYPE tp_defined_names IS TABLE OF tp_defined_name INDEX BY PLS_INTEGER;
  TYPE tp_book IS RECORD(
     sheets        tp_sheets
    ,strings       tp_strings
    ,str_ind       tp_str_ind
    ,str_cnt       PLS_INTEGER := 0
    ,fonts         tp_fonts
    ,fills         tp_fills
    ,borders       tp_borders
    ,numFmts       tp_numFmts
    ,cellXfs       tp_cellXfs
    ,numFmtIndexes tp_numFmtIndexes
    ,defined_names tp_defined_names);
  workbook tp_book;
  --
  PROCEDURE blob2file(p_blob      BLOB
                     ,p_directory VARCHAR2 := 'MY_DIR'
                     ,p_filename  VARCHAR2 := 'my.xlsx') IS
    t_fh  utl_file.file_type;
    t_len PLS_INTEGER := 32767;
  BEGIN
    t_fh := utl_file.fopen(p_directory
                          ,p_filename
                          ,'wb');
    FOR i IN 0 .. trunc((dbms_lob.getlength(p_blob) - 1) / t_len)
    LOOP
      utl_file.put_raw(t_fh
                      ,dbms_lob.substr(p_blob
                                      ,t_len
                                      ,i * t_len + 1));
    END LOOP;
    utl_file.fclose(t_fh);
  END;
  --
  FUNCTION raw2num(p_raw RAW
                  ,p_len INTEGER
                  ,p_pos INTEGER) RETURN NUMBER IS
  BEGIN
    RETURN utl_raw.cast_to_binary_integer(utl_raw.substr(p_raw
                                                        ,p_pos
                                                        ,p_len)
                                         ,utl_raw.little_endian);
  END;
  --
  FUNCTION little_endian(p_big   NUMBER
                        ,p_bytes PLS_INTEGER := 4) RETURN RAW IS
  BEGIN
    RETURN utl_raw.substr(utl_raw.cast_from_binary_integer(p_big
                                                          ,utl_raw.little_endian)
                         ,1
                         ,p_bytes);
  END;
  --
  FUNCTION blob2num(p_blob BLOB
                   ,p_len  INTEGER
                   ,p_pos  INTEGER) RETURN NUMBER IS
  BEGIN
    RETURN utl_raw.cast_to_binary_integer(dbms_lob.substr(p_blob
                                                         ,p_len
                                                         ,p_pos)
                                         ,utl_raw.little_endian);
  END;
  --
  PROCEDURE add1file(p_zipped_blob IN OUT BLOB
                    ,p_name        VARCHAR2
                    ,p_content     BLOB) IS
    t_now        DATE;
    t_blob       BLOB;
    t_len        INTEGER;
    t_clen       INTEGER;
    t_crc32      RAW(4) := hextoraw('00000000');
    t_compressed BOOLEAN := FALSE;
    t_name       RAW(32767);
  BEGIN
    t_now := SYSDATE;
    t_len := NVL(dbms_lob.getlength(p_content)
                ,0);
    IF t_len > 0
    THEN
      t_blob       := utl_compress.lz_compress(p_content);
      t_clen       := dbms_lob.getlength(t_blob) - 18;
      t_compressed := t_clen < t_len;
      t_crc32      := dbms_lob.substr(t_blob
                                     ,4
                                     ,t_clen + 11);
    END IF;
    IF NOT t_compressed
    THEN
      t_clen := t_len;
      t_blob := p_content;
    END IF;
    IF p_zipped_blob IS NULL
    THEN
      dbms_lob.createtemporary(p_zipped_blob
                              ,TRUE);
    END IF;
    t_name := utl_i18n.string_to_raw(p_name
                                    ,'AL32UTF8');
    dbms_lob.append(p_zipped_blob
                   ,utl_raw.concat(c_LOCAL_FILE_HEADER -- Local file header signature
                                  ,hextoraw('1400') -- version 2.0
                                  ,CASE WHEN
                                   t_name =
                                   utl_i18n.string_to_raw(p_name
                                                         ,'US8PC437') THEN
                                   hextoraw('0000') -- no General purpose bits
                                   ELSE hextoraw('0008') -- set Language encoding flag (EFS)
                                   END
                                  ,CASE WHEN t_compressed THEN
                                   hextoraw('0800') -- deflate
                                   ELSE hextoraw('0000') -- stored
                                   END
                                  ,little_endian(to_number(to_char(t_now
                                                                  ,'ss')) / 2 +
                                                 to_number(to_char(t_now
                                                                  ,'mi')) * 32 +
                                                 to_number(to_char(t_now
                                                                  ,'hh24')) * 2048
                                                ,2) -- File last modification time
                                  ,little_endian(to_number(to_char(t_now
                                                                  ,'dd')) +
                                                 to_number(to_char(t_now
                                                                  ,'mm')) * 32 +
                                                 (to_number(to_char(t_now
                                                                   ,'yyyy')) - 1980) * 512
                                                ,2) -- File last modification date
                                  ,t_crc32 -- CRC-32
                                  ,little_endian(t_clen) -- compressed size
                                  ,little_endian(t_len) -- uncompressed size
                                  ,little_endian(utl_raw.length(t_name)
                                                ,2) -- File name length
                                  ,hextoraw('0000') -- Extra field length
                                  ,t_name -- File name
                                   ));
    IF t_compressed
    THEN
      dbms_lob.copy(p_zipped_blob
                   ,t_blob
                   ,t_clen
                   ,dbms_lob.getlength(p_zipped_blob) + 1
                   ,11); -- compressed content
    ELSIF t_clen > 0
    THEN
      dbms_lob.copy(p_zipped_blob
                   ,t_blob
                   ,t_clen
                   ,dbms_lob.getlength(p_zipped_blob) + 1
                   ,1); --  content
    END IF;
    IF dbms_lob.istemporary(t_blob) = 1
    THEN
      dbms_lob.freetemporary(t_blob);
    END IF;
  END;
  --
  PROCEDURE finish_zip(p_zipped_blob IN OUT BLOB) IS
    t_cnt             PLS_INTEGER := 0;
    t_offs            INTEGER;
    t_offs_dir_header INTEGER;
    t_offs_end_header INTEGER;
    t_comment         RAW(32767) := utl_raw.cast_to_raw('Implementation by Andrzej Raczkowski');
  BEGIN
    t_offs_dir_header := dbms_lob.getlength(p_zipped_blob);
    t_offs            := 1;
    WHILE dbms_lob.substr(p_zipped_blob
                         ,utl_raw.length(c_LOCAL_FILE_HEADER)
                         ,t_offs) = c_LOCAL_FILE_HEADER
    LOOP
      t_cnt := t_cnt + 1;
      dbms_lob.append(p_zipped_blob
                     ,utl_raw.concat(hextoraw('504B0102') -- Central directory file header signature
                                    ,hextoraw('1400') -- version 2.0
                                    ,dbms_lob.substr(p_zipped_blob
                                                    ,26
                                                    ,t_offs + 4)
                                    ,hextoraw('0000') -- File comment length
                                    ,hextoraw('0000') -- Disk number where file starts
                                    ,hextoraw('0000') -- Internal file attributes =>
                                     --     0000 binary file
                                     --     0100 (ascii)text file
                                    ,CASE WHEN dbms_lob.substr(p_zipped_blob
                                                    ,1
                                                    ,t_offs + 30 +
                                                     blob2num(p_zipped_blob
                                                             ,2
                                                             ,t_offs + 26) - 1) IN
                                     (hextoraw('2F') -- /
                                    ,hextoraw('5C') -- \
                                      ) THEN hextoraw('10000000') -- a directory/folder
                                     ELSE hextoraw('2000B681') -- a file
                                     END -- External file attributes
                                    ,little_endian(t_offs - 1) -- Relative offset of local file header
                                    ,dbms_lob.substr(p_zipped_blob
                                                    ,blob2num(p_zipped_blob
                                                             ,2
                                                             ,t_offs + 26)
                                                    ,t_offs + 30) -- File name
                                     ));
      t_offs := t_offs + 30 + blob2num(p_zipped_blob
                                      ,4
                                      ,t_offs + 18) -- compressed size
                + blob2num(p_zipped_blob
                          ,2
                          ,t_offs + 26) -- File name length
                + blob2num(p_zipped_blob
                          ,2
                          ,t_offs + 28); -- Extra field length
    END LOOP;
    t_offs_end_header := dbms_lob.getlength(p_zipped_blob);
    dbms_lob.append(p_zipped_blob
                   ,utl_raw.concat(c_END_OF_CENTRAL_DIRECTORY -- End of central directory signature
                                  ,hextoraw('0000') -- Number of this disk
                                  ,hextoraw('0000') -- Disk where central directory starts
                                  ,little_endian(t_cnt
                                                ,2) -- Number of central directory records on this disk
                                  ,little_endian(t_cnt
                                                ,2) -- Total number of central directory records
                                  ,little_endian(t_offs_end_header -
                                                 t_offs_dir_header) -- Size of central directory
                                  ,little_endian(t_offs_dir_header) -- Offset of start of central directory, relative to start of archive
                                  ,little_endian(NVL(utl_raw.length(t_comment)
                                                    ,0)
                                                ,2) -- ZIP file comment length
                                  ,t_comment));
  END;
  --
  FUNCTION alfan_col(p_col PLS_INTEGER) RETURN VARCHAR2 IS
  BEGIN
    RETURN CASE WHEN p_col > 702 THEN chr(64 + trunc((p_col - 27) / 676)) || chr(65 +
                                                                                 MOD(trunc((p_col - 1) / 26) - 1
                                                                                    ,26)) || chr(65 +
                                                                                                 MOD(p_col - 1
                                                                                                    ,26)) WHEN p_col > 26 THEN chr(64 +
                                                                                                                                   trunc((p_col - 1) / 26)) || chr(65 +
                                                                                                                                                                   MOD(p_col - 1
                                                                                                                                                                      ,26)) ELSE chr(64 +
                                                                                                                                                                                     p_col) END;
  END;
  --
  FUNCTION col_alfan(p_col VARCHAR2) RETURN PLS_INTEGER IS
  BEGIN
    RETURN ascii(substr(p_col
                       ,-1)) - 64 + NVL((ascii(substr(p_col
                                                     ,-2
                                                     ,1)) - 64) * 26
                                       ,0) + NVL((ascii(substr(p_col
                                                              ,-3
                                                              ,1)) - 64) * 676
                                                ,0);
  END;
  --
  PROCEDURE clear_workbook IS
    t_row_ind PLS_INTEGER;
  BEGIN
    FOR s IN 1 .. workbook.sheets.count()
    LOOP
      t_row_ind := workbook.sheets(s).rows.first();
      WHILE t_row_ind IS NOT NULL
      LOOP
        workbook.sheets(s).rows(t_row_ind).delete();
        t_row_ind := workbook.sheets(s).rows.next(t_row_ind);
      END LOOP;
      workbook.sheets(s).rows.delete();
      workbook.sheets(s).widths.delete();
      workbook.sheets(s).autofilters.delete();
      workbook.sheets(s).hyperlinks.delete();
      workbook.sheets(s).col_fmts.delete();
      workbook.sheets(s).row_fmts.delete();
      workbook.sheets(s).comments.delete();
      workbook.sheets(s).mergecells.delete();
      workbook.sheets(s).validations.delete();
    END LOOP;
    workbook.strings.delete();
    workbook.str_ind.delete();
    workbook.fonts.delete();
    workbook.fills.delete();
    workbook.borders.delete();
    workbook.numFmts.delete();
    workbook.cellXfs.delete();
    workbook.defined_names.delete();
    workbook := NULL;
  END;
  --
  PROCEDURE new_sheet(p_sheetname VARCHAR2 := NULL) IS
    t_nr  PLS_INTEGER := workbook.sheets.count() + 1;
    t_ind PLS_INTEGER;
  BEGIN
    workbook.sheets(t_nr).name := NVL(dbms_xmlgen.convert(translate(p_sheetname
                                                                   ,'a/\[]*:?'
                                                                   ,'a'))
                                     ,'Sheet' || t_nr);
    IF workbook.strings.count() = 0
    THEN
      workbook.str_cnt := 0;
    END IF;
    IF workbook.fonts.count() = 0
    THEN
      t_ind := get_font('Calibri');
    END IF;
    IF workbook.fills.count() = 0
    THEN
      t_ind := get_fill('none');
      t_ind := get_fill('gray125');
    END IF;
    IF workbook.borders.count() = 0
    THEN
      t_ind := get_border(''
                         ,''
                         ,''
                         ,'');
    END IF;
  END;
  --
  PROCEDURE set_col_width(p_sheet  PLS_INTEGER
                         ,p_col    PLS_INTEGER
                         ,p_format VARCHAR2) IS
    t_width  NUMBER;
    t_nr_chr PLS_INTEGER;
  BEGIN
    IF p_format IS NULL
    THEN
      RETURN;
    END IF;
    IF instr(p_format
            ,';') > 0
    THEN
      t_nr_chr := length(translate(substr(p_format
                                         ,1
                                         ,instr(p_format
                                               ,';') - 1)
                                  ,'a\"'
                                  ,'a'));
    ELSE
      t_nr_chr := length(translate(p_format
                                  ,'a\"'
                                  ,'a'));
    END IF;
    t_width := trunc((t_nr_chr * 7 + 5) / 7 * 256) / 256; -- assume default 11 point Calibri
    IF workbook.sheets(p_sheet).widths.exists(p_col)
    THEN
      workbook.sheets(p_sheet).widths(p_col) := greatest(workbook.sheets(p_sheet)
                                                         .widths(p_col)
                                                        ,t_width);
    ELSE
      workbook.sheets(p_sheet).widths(p_col) := greatest(t_width
                                                        ,8.43);
    END IF;
  END;
  --
  FUNCTION OraFmt2Excel(p_format VARCHAR2 := NULL) RETURN VARCHAR2 IS
    t_format VARCHAR2(1000) := substr(p_format
                                     ,1
                                     ,1000);
  BEGIN
    t_format := REPLACE(REPLACE(t_format
                               ,'hh24'
                               ,'hh')
                       ,'hh12'
                       ,'hh');
    t_format := REPLACE(t_format
                       ,'mi'
                       ,'mm');
    t_format := REPLACE(REPLACE(REPLACE(t_format
                                       ,'AM'
                                       ,'~~')
                               ,'PM'
                               ,'~~')
                       ,'~~'
                       ,'AM/PM');
    t_format := REPLACE(REPLACE(REPLACE(t_format
                                       ,'am'
                                       ,'~~')
                               ,'pm'
                               ,'~~')
                       ,'~~'
                       ,'AM/PM');
    t_format := REPLACE(REPLACE(t_format
                               ,'day'
                               ,'DAY')
                       ,'DAY'
                       ,'dddd');
    t_format := REPLACE(REPLACE(t_format
                               ,'dy'
                               ,'DY')
                       ,'DAY'
                       ,'ddd');
    t_format := REPLACE(REPLACE(t_format
                               ,'RR'
                               ,'RR')
                       ,'RR'
                       ,'YY');
    t_format := REPLACE(REPLACE(t_format
                               ,'month'
                               ,'MONTH')
                       ,'MONTH'
                       ,'mmmm');
    t_format := REPLACE(REPLACE(t_format
                               ,'mon'
                               ,'MON')
                       ,'MON'
                       ,'mmm');
    RETURN t_format;
  END;
  --
  FUNCTION get_numFmt(p_format VARCHAR2 := NULL) RETURN PLS_INTEGER IS
    t_cnt      PLS_INTEGER;
    t_numFmtId PLS_INTEGER;
  BEGIN
    IF p_format IS NULL
    THEN
      RETURN 0;
    END IF;
    t_cnt := workbook.numFmts.count();
    FOR i IN 1 .. t_cnt
    LOOP
      IF workbook.numFmts(i).formatCode = p_format
      THEN
        t_numFmtId := workbook.numFmts(i).numFmtId;
        EXIT;
      END IF;
    END LOOP;
    IF t_numFmtId IS NULL
    THEN
      t_numFmtId := CASE
                      WHEN t_cnt = 0 THEN
                       164
                      ELSE
                       workbook.numFmts(t_cnt).numFmtId + 1
                    END;
      t_cnt := t_cnt + 1;
      workbook.numFmts(t_cnt).numFmtId := t_numFmtId;
      workbook.numFmts(t_cnt).formatCode := p_format;
      workbook.numFmtIndexes(t_numFmtId) := t_cnt;
    END IF;
    RETURN t_numFmtId;
  END;
  --
  FUNCTION get_font(p_name      VARCHAR2
                   ,p_family    PLS_INTEGER := 2
                   ,p_fontsize  NUMBER := 11
                   ,p_theme     PLS_INTEGER := 1
                   ,p_underline BOOLEAN := FALSE
                   ,p_italic    BOOLEAN := FALSE
                   ,p_bold      BOOLEAN := FALSE
                   ,p_rgb       VARCHAR2 := NULL -- this is a hex ALPHA Red Green Blue value
                    ) RETURN PLS_INTEGER IS
    t_ind PLS_INTEGER;
  BEGIN
    IF workbook.fonts.count() > 0
    THEN
      FOR f IN 0 .. workbook.fonts.count() - 1
      LOOP
        IF (workbook.fonts(f)
           .name = p_name AND workbook.fonts(f).family = p_family AND workbook.fonts(f)
           .fontsize = p_fontsize AND workbook.fonts(f).theme = p_theme AND workbook.fonts(f)
           .underline = p_underline AND workbook.fonts(f).italic = p_italic AND workbook.fonts(f)
           .bold = p_bold AND
            (workbook.fonts(f)
            .rgb = p_rgb OR (workbook.fonts(f).rgb IS NULL AND p_rgb IS NULL)))
        THEN
          RETURN f;
        END IF;
      END LOOP;
    END IF;
    t_ind := workbook.fonts.count();
    workbook.fonts(t_ind).name := p_name;
    workbook.fonts(t_ind).family := p_family;
    workbook.fonts(t_ind).fontsize := p_fontsize;
    workbook.fonts(t_ind).theme := p_theme;
    workbook.fonts(t_ind).underline := p_underline;
    workbook.fonts(t_ind).italic := p_italic;
    workbook.fonts(t_ind).bold := p_bold;
    workbook.fonts(t_ind).rgb := p_rgb;
    RETURN t_ind;
  END;
  --
  FUNCTION get_fill(p_patternType VARCHAR2
                   ,p_fgRGB       VARCHAR2 := NULL) RETURN PLS_INTEGER IS
    t_ind PLS_INTEGER;
  BEGIN
    IF workbook.fills.count() > 0
    THEN
      FOR f IN 0 .. workbook.fills.count() - 1
      LOOP
        IF (workbook.fills(f)
           .patternType = p_patternType AND
            NVL(workbook.fills(f).fgRGB
               ,'x') = NVL(upper(p_fgRGB)
                          ,'x'))
        THEN
          RETURN f;
        END IF;
      END LOOP;
    END IF;
    t_ind := workbook.fills.count();
    workbook.fills(t_ind).patternType := p_patternType;
    workbook.fills(t_ind).fgRGB := upper(p_fgRGB);
    RETURN t_ind;
  END;
  --
  FUNCTION get_border(p_top    VARCHAR2 := 'thin'
                     ,p_bottom VARCHAR2 := 'thin'
                     ,p_left   VARCHAR2 := 'thin'
                     ,p_right  VARCHAR2 := 'thin') RETURN PLS_INTEGER IS
    t_ind PLS_INTEGER;
  BEGIN
    IF workbook.borders.count() > 0
    THEN
      FOR b IN 0 .. workbook.borders.count() - 1
      LOOP
        IF (NVL(workbook.borders(b).top
               ,'x') = NVL(p_top
                           ,'x') AND
           NVL(workbook.borders(b).bottom
               ,'x') = NVL(p_bottom
                           ,'x') AND
           NVL(workbook.borders(b).left
               ,'x') = NVL(p_left
                           ,'x') AND
           NVL(workbook.borders(b).right
               ,'x') = NVL(p_right
                           ,'x'))
        THEN
          RETURN b;
        END IF;
      END LOOP;
    END IF;
    t_ind := workbook.borders.count();
    workbook.borders(t_ind).top := p_top;
    workbook.borders(t_ind).bottom := p_bottom;
    workbook.borders(t_ind).left := p_left;
    workbook.borders(t_ind).right := p_right;
    RETURN t_ind;
  END;
  --
  FUNCTION get_alignment(p_vertical   VARCHAR2 := NULL
                        ,p_horizontal VARCHAR2 := NULL
                        ,p_wrapText   BOOLEAN := NULL) RETURN tp_alignment IS
    t_rv tp_alignment;
  BEGIN
    t_rv.vertical   := p_vertical;
    t_rv.horizontal := p_horizontal;
    t_rv.wrapText   := p_wrapText;
    RETURN t_rv;
  END;
  --
  FUNCTION get_XfId(p_sheet     PLS_INTEGER
                   ,p_col       PLS_INTEGER
                   ,p_row       PLS_INTEGER
                   ,p_numFmtId  PLS_INTEGER := NULL
                   ,p_fontId    PLS_INTEGER := NULL
                   ,p_fillId    PLS_INTEGER := NULL
                   ,p_borderId  PLS_INTEGER := NULL
                   ,p_alignment tp_alignment := NULL) RETURN VARCHAR2 IS
    t_cnt    PLS_INTEGER;
    t_XfId   PLS_INTEGER;
    t_XF     tp_XF_fmt;
    t_col_XF tp_XF_fmt;
    t_row_XF tp_XF_fmt;
  BEGIN
    IF workbook.sheets(p_sheet).col_fmts.exists(p_col)
    THEN
      t_col_XF := workbook.sheets(p_sheet).col_fmts(p_col);
    END IF;
    IF workbook.sheets(p_sheet).row_fmts.exists(p_row)
    THEN
      t_row_XF := workbook.sheets(p_sheet).row_fmts(p_row);
    END IF;
    t_XF.numFmtId  := coalesce(p_numFmtId
                              ,t_col_XF.numFmtId
                              ,t_row_XF.numFmtId
                              ,0);
    t_XF.fontId    := coalesce(p_fontId
                              ,t_col_XF.fontId
                              ,t_row_XF.fontId
                              ,0);
    t_XF.fillId    := coalesce(p_fillId
                              ,t_col_XF.fillId
                              ,t_row_XF.fillId
                              ,0);
    t_XF.borderId  := coalesce(p_borderId
                              ,t_col_XF.borderId
                              ,t_row_XF.borderId
                              ,0);
    t_XF.alignment := coalesce(p_alignment
                              ,t_col_XF.alignment
                              ,t_row_XF.alignment);
    IF (t_XF.numFmtId + t_XF.fontId + t_XF.fillId + t_XF.borderId = 0 AND
       t_XF.alignment.vertical IS NULL AND t_XF.alignment.horizontal IS NULL AND
       NOT NVL(t_XF.alignment.wrapText
               ,FALSE))
    THEN
      RETURN '';
    END IF;
    IF t_XF.numFmtId > 0
    THEN
      set_col_width(p_sheet
                   ,p_col
                   ,workbook.numFmts(workbook.numFmtIndexes(t_XF.numFmtId))
                    .formatCode);
    END IF;
    t_cnt := workbook.cellXfs.count();
    FOR i IN 1 .. t_cnt
    LOOP
      IF (workbook.cellXfs(i)
         .numFmtId = t_XF.numFmtId AND workbook.cellXfs(i)
         .fontId = t_XF.fontId AND workbook.cellXfs(i).fillId = t_XF.fillId AND workbook.cellXfs(i)
         .borderId = t_XF.borderId AND
          NVL(workbook.cellXfs(i).alignment.vertical
             ,'x') = NVL(t_XF.alignment.vertical
                        ,'x') AND
          NVL(workbook.cellXfs(i).alignment.horizontal
             ,'x') = NVL(t_XF.alignment.horizontal
                        ,'x') AND
          NVL(workbook.cellXfs(i).alignment.wrapText
             ,FALSE) = NVL(t_XF.alignment.wrapText
                          ,FALSE))
      THEN
        t_XfId := i;
        EXIT;
      END IF;
    END LOOP;
    IF t_XfId IS NULL
    THEN
      t_cnt := t_cnt + 1;
      t_XfId := t_cnt;
      workbook.cellXfs(t_cnt) := t_XF;
    END IF;
    RETURN 's="' || t_XfId || '"';
  END;
  --
  PROCEDURE cell(p_col       PLS_INTEGER
                ,p_row       PLS_INTEGER
                ,p_value     NUMBER
                ,p_numFmtId  PLS_INTEGER := NULL
                ,p_fontId    PLS_INTEGER := NULL
                ,p_fillId    PLS_INTEGER := NULL
                ,p_borderId  PLS_INTEGER := NULL
                ,p_alignment tp_alignment := NULL
                ,p_sheet     PLS_INTEGER := NULL) IS
    t_sheet PLS_INTEGER := NVL(p_sheet
                              ,workbook.sheets.count());
  BEGIN
    workbook.sheets(t_sheet).rows(p_row)(p_col).value := p_value;
    workbook.sheets(t_sheet).rows(p_row)(p_col).style := NULL;
    workbook.sheets(t_sheet).rows(p_row)(p_col).style := get_XfId(t_sheet
                                                                 ,p_col
                                                                 ,p_row
                                                                 ,p_numFmtId
                                                                 ,p_fontId
                                                                 ,p_fillId
                                                                 ,p_borderId
                                                                 ,p_alignment);
  END;
  --
  FUNCTION add_string(p_string VARCHAR2) RETURN PLS_INTEGER IS
    t_cnt PLS_INTEGER;
  BEGIN
    IF workbook.strings.exists(p_string)
    THEN
      t_cnt := workbook.strings(p_string);
    ELSE
      t_cnt := workbook.strings.count();
      workbook.str_ind(t_cnt) := p_string;
      workbook.strings(NVL(p_string, '')) := t_cnt;
    END IF;
    workbook.str_cnt := workbook.str_cnt + 1;
    RETURN t_cnt;
  END;
  --
  PROCEDURE cell(p_col       PLS_INTEGER
                ,p_row       PLS_INTEGER
                ,p_value     VARCHAR2
                ,p_numFmtId  PLS_INTEGER := NULL
                ,p_fontId    PLS_INTEGER := NULL
                ,p_fillId    PLS_INTEGER := NULL
                ,p_borderId  PLS_INTEGER := NULL
                ,p_alignment tp_alignment := NULL
                ,p_sheet     PLS_INTEGER := NULL) IS
    t_sheet     PLS_INTEGER := NVL(p_sheet
                                  ,workbook.sheets.count());
    t_alignment tp_alignment := p_alignment;
  BEGIN
    workbook.sheets(t_sheet).rows(p_row)(p_col).value := add_string(p_value);
    IF t_alignment.wrapText IS NULL
       AND instr(p_value
                ,chr(13)) > 0
    THEN
      t_alignment.wrapText := TRUE;
    END IF;
    workbook.sheets(t_sheet).rows(p_row)(p_col).style := 't="s" ' ||
                                                         get_XfId(t_sheet
                                                                 ,p_col
                                                                 ,p_row
                                                                 ,p_numFmtId
                                                                 ,p_fontId
                                                                 ,p_fillId
                                                                 ,p_borderId
                                                                 ,t_alignment);
  END;
  --
  PROCEDURE cell(p_col       PLS_INTEGER
                ,p_row       PLS_INTEGER
                ,p_value     DATE
                ,p_numFmtId  PLS_INTEGER := NULL
                ,p_fontId    PLS_INTEGER := NULL
                ,p_fillId    PLS_INTEGER := NULL
                ,p_borderId  PLS_INTEGER := NULL
                ,p_alignment tp_alignment := NULL
                ,p_sheet     PLS_INTEGER := NULL) IS
    t_numFmtId PLS_INTEGER := p_numFmtId;
    t_sheet    PLS_INTEGER := NVL(p_sheet
                                 ,workbook.sheets.count());
  BEGIN
    workbook.sheets(t_sheet).rows(p_row)(p_col).value := p_value -
                                                         to_date('01-01-1904'
                                                                ,'DD-MM-YYYY');
    IF t_numFmtId IS NULL
       AND NOT (workbook.sheets(t_sheet).col_fmts.exists(p_col) AND workbook.sheets(t_sheet).col_fmts(p_col)
        .numFmtId IS NOT NULL)
       AND NOT (workbook.sheets(t_sheet).row_fmts.exists(p_row) AND workbook.sheets(t_sheet).row_fmts(p_row)
        .numFmtId IS NOT NULL)
    THEN
      t_numFmtId := get_numFmt('dd/mm/yyyy');
    END IF;
    workbook.sheets(t_sheet).rows(p_row)(p_col).style := get_XfId(t_sheet
                                                                 ,p_col
                                                                 ,p_row
                                                                 ,t_numFmtId
                                                                 ,p_fontId
                                                                 ,p_fillId
                                                                 ,p_borderId
                                                                 ,p_alignment);
  END;
  --
  PROCEDURE hyperlink(p_col   PLS_INTEGER
                     ,p_row   PLS_INTEGER
                     ,p_url   VARCHAR2
                     ,p_value VARCHAR2 := NULL
                     ,p_sheet PLS_INTEGER := NULL) IS
    t_ind   PLS_INTEGER;
    t_sheet PLS_INTEGER := NVL(p_sheet
                              ,workbook.sheets.count());
  BEGIN
    workbook.sheets(t_sheet).rows(p_row)(p_col).value := add_string(NVL(p_value
                                                                       ,p_url));
    workbook.sheets(t_sheet).rows(p_row)(p_col).style := 't="s" ' ||
                                                         get_XfId(t_sheet
                                                                 ,p_col
                                                                 ,p_row
                                                                 ,''
                                                                 ,get_font('Calibri'
                                                                          ,p_theme     => 10
                                                                          ,p_underline => TRUE));
    t_ind := workbook.sheets(t_sheet).hyperlinks.count() + 1;
    workbook.sheets(t_sheet).hyperlinks(t_ind).cell := alfan_col(p_col) ||
                                                       p_row;
    workbook.sheets(t_sheet).hyperlinks(t_ind).url := p_url;
  END;
  --
  PROCEDURE COMMENT(p_col    PLS_INTEGER
                   ,p_row    PLS_INTEGER
                   ,p_text   VARCHAR2
                   ,p_author VARCHAR2 := NULL
                   ,p_width  PLS_INTEGER := 150
                   ,p_height PLS_INTEGER := 100
                   ,p_sheet  PLS_INTEGER := NULL) IS
    t_ind   PLS_INTEGER;
    t_sheet PLS_INTEGER := NVL(p_sheet
                              ,workbook.sheets.count());
  BEGIN
    t_ind := workbook.sheets(t_sheet).comments.count() + 1;
    workbook.sheets(t_sheet).comments(t_ind).row := p_row;
    workbook.sheets(t_sheet).comments(t_ind).column := p_col;
    workbook.sheets(t_sheet).comments(t_ind).text := dbms_xmlgen.convert(p_text);
    workbook.sheets(t_sheet).comments(t_ind).author := dbms_xmlgen.convert(p_author);
    workbook.sheets(t_sheet).comments(t_ind).width := p_width;
    workbook.sheets(t_sheet).comments(t_ind).height := p_height;
  END;
  --
  PROCEDURE mergecells(p_tl_col PLS_INTEGER -- top left
                      ,p_tl_row PLS_INTEGER
                      ,p_br_col PLS_INTEGER -- bottom right
                      ,p_br_row PLS_INTEGER
                      ,p_sheet  PLS_INTEGER := NULL) IS
    t_ind   PLS_INTEGER;
    t_sheet PLS_INTEGER := NVL(p_sheet
                              ,workbook.sheets.count());
  BEGIN
    t_ind := workbook.sheets(t_sheet).mergecells.count() + 1;
    workbook.sheets(t_sheet).mergecells(t_ind) := alfan_col(p_tl_col) ||
                                                  p_tl_row || ':' ||
                                                  alfan_col(p_br_col) ||
                                                  p_br_row;
  END;
  --
  PROCEDURE add_validation(p_type        VARCHAR2
                          ,p_sqref       VARCHAR2
                          ,p_style       VARCHAR2 := 'stop' -- stop, warning, information
                          ,p_formula1    VARCHAR2 := NULL
                          ,p_formula2    VARCHAR2 := NULL
                          ,p_title       VARCHAR2 := NULL
                          ,p_prompt      VARCHAR := NULL
                          ,p_show_error  BOOLEAN := FALSE
                          ,p_error_title VARCHAR2 := NULL
                          ,p_error_txt   VARCHAR2 := NULL
                          ,p_sheet       PLS_INTEGER := NULL) IS
    t_ind   PLS_INTEGER;
    t_sheet PLS_INTEGER := NVL(p_sheet
                              ,workbook.sheets.count());
  BEGIN
    t_ind := workbook.sheets(t_sheet).validations.count() + 1;
    workbook.sheets(t_sheet).validations(t_ind).type := p_type;
    workbook.sheets(t_sheet).validations(t_ind).errorstyle := p_style;
    workbook.sheets(t_sheet).validations(t_ind).sqref := p_sqref;
    workbook.sheets(t_sheet).validations(t_ind).formula1 := p_formula1;
    workbook.sheets(t_sheet).validations(t_ind).error_title := p_error_title;
    workbook.sheets(t_sheet).validations(t_ind).error_txt := p_error_txt;
    workbook.sheets(t_sheet).validations(t_ind).title := p_title;
    workbook.sheets(t_sheet).validations(t_ind).prompt := p_prompt;
    workbook.sheets(t_sheet).validations(t_ind).showerrormessage := p_show_error;
  END;
  --
  PROCEDURE list_validation(p_sqref_col   PLS_INTEGER
                           ,p_sqref_row   PLS_INTEGER
                           ,p_tl_col      PLS_INTEGER -- top left
                           ,p_tl_row      PLS_INTEGER
                           ,p_br_col      PLS_INTEGER -- bottom right
                           ,p_br_row      PLS_INTEGER
                           ,p_style       VARCHAR2 := 'stop' -- stop, warning, information
                           ,p_title       VARCHAR2 := NULL
                           ,p_prompt      VARCHAR := NULL
                           ,p_show_error  BOOLEAN := FALSE
                           ,p_error_title VARCHAR2 := NULL
                           ,p_error_txt   VARCHAR2 := NULL
                           ,p_sheet       PLS_INTEGER := NULL) IS
  BEGIN
    add_validation('list'
                  ,alfan_col(p_sqref_col) || p_sqref_row
                  ,p_style => lower(p_style)
                  ,p_formula1 => '$' || alfan_col(p_tl_col) || '$' ||
                                 p_tl_row || ':$' || alfan_col(p_br_col) || '$' ||
                                 p_br_row
                  ,p_title => p_title
                  ,p_prompt => p_prompt
                  ,p_show_error => p_show_error
                  ,p_error_title => p_error_title
                  ,p_error_txt => p_error_txt
                  ,p_sheet => p_sheet);
  END;
  --
  PROCEDURE list_validation(p_sqref_col    PLS_INTEGER
                           ,p_sqref_row    PLS_INTEGER
                           ,p_defined_name VARCHAR2
                           ,p_style        VARCHAR2 := 'stop' -- stop, warning, information
                           ,p_title        VARCHAR2 := NULL
                           ,p_prompt       VARCHAR := NULL
                           ,p_show_error   BOOLEAN := FALSE
                           ,p_error_title  VARCHAR2 := NULL
                           ,p_error_txt    VARCHAR2 := NULL
                           ,p_sheet        PLS_INTEGER := NULL) IS
  BEGIN
    add_validation('list'
                  ,alfan_col(p_sqref_col) || p_sqref_row
                  ,p_style => lower(p_style)
                  ,p_formula1 => p_defined_name
                  ,p_title => p_title
                  ,p_prompt => p_prompt
                  ,p_show_error => p_show_error
                  ,p_error_title => p_error_title
                  ,p_error_txt => p_error_txt
                  ,p_sheet => p_sheet);
  END;
  --
  PROCEDURE defined_name(p_tl_col     PLS_INTEGER -- top left
                        ,p_tl_row     PLS_INTEGER
                        ,p_br_col     PLS_INTEGER -- bottom right
                        ,p_br_row     PLS_INTEGER
                        ,p_name       VARCHAR2
                        ,p_sheet      PLS_INTEGER := NULL
                        ,p_localsheet PLS_INTEGER := NULL) IS
    t_ind   PLS_INTEGER;
    t_sheet PLS_INTEGER := NVL(p_sheet
                              ,workbook.sheets.count());
  BEGIN
    t_ind := workbook.defined_names.count() + 1;
    workbook.defined_names(t_ind).name := p_name;
    workbook.defined_names(t_ind).ref := 'Sheet' || t_sheet || '!$' ||
                                         alfan_col(p_tl_col) || '$' ||
                                         p_tl_row || ':$' ||
                                         alfan_col(p_br_col) || '$' ||
                                         p_br_row;
    workbook.defined_names(t_ind).sheet := p_localsheet;
  END;
  --
  PROCEDURE set_column_width(p_col   PLS_INTEGER
                            ,p_width NUMBER
                            ,p_sheet PLS_INTEGER := NULL) IS
  BEGIN
    workbook.sheets(NVL(p_sheet, workbook.sheets.count())).widths(p_col) := p_width;
  END;
  --
  PROCEDURE set_column(p_col       PLS_INTEGER
                      ,p_numFmtId  PLS_INTEGER := NULL
                      ,p_fontId    PLS_INTEGER := NULL
                      ,p_fillId    PLS_INTEGER := NULL
                      ,p_borderId  PLS_INTEGER := NULL
                      ,p_alignment tp_alignment := NULL
                      ,p_sheet     PLS_INTEGER := NULL) IS
    t_sheet PLS_INTEGER := NVL(p_sheet
                              ,workbook.sheets.count());
  BEGIN
    workbook.sheets(t_sheet).col_fmts(p_col).numFmtId := p_numFmtId;
    workbook.sheets(t_sheet).col_fmts(p_col).fontId := p_fontId;
    workbook.sheets(t_sheet).col_fmts(p_col).fillId := p_fillId;
    workbook.sheets(t_sheet).col_fmts(p_col).borderId := p_borderId;
    workbook.sheets(t_sheet).col_fmts(p_col).alignment := p_alignment;
  END;
  --
  PROCEDURE set_row(p_row       PLS_INTEGER
                   ,p_numFmtId  PLS_INTEGER := NULL
                   ,p_fontId    PLS_INTEGER := NULL
                   ,p_fillId    PLS_INTEGER := NULL
                   ,p_borderId  PLS_INTEGER := NULL
                   ,p_alignment tp_alignment := NULL
                   ,p_sheet     PLS_INTEGER := NULL) IS
    t_sheet PLS_INTEGER := NVL(p_sheet
                              ,workbook.sheets.count());
  BEGIN
    workbook.sheets(t_sheet).row_fmts(p_row).numFmtId := p_numFmtId;
    workbook.sheets(t_sheet).row_fmts(p_row).fontId := p_fontId;
    workbook.sheets(t_sheet).row_fmts(p_row).fillId := p_fillId;
    workbook.sheets(t_sheet).row_fmts(p_row).borderId := p_borderId;
    workbook.sheets(t_sheet).row_fmts(p_row).alignment := p_alignment;
  END;
  --
  PROCEDURE freeze_rows(p_nr_rows PLS_INTEGER := 1
                       ,p_sheet   PLS_INTEGER := NULL) IS
    t_sheet PLS_INTEGER := NVL(p_sheet
                              ,workbook.sheets.count());
  BEGIN
    workbook.sheets(t_sheet).freeze_cols := NULL;
    workbook.sheets(t_sheet).freeze_rows := p_nr_rows;
  END;
  --
  PROCEDURE freeze_cols(p_nr_cols PLS_INTEGER := 1
                       ,p_sheet   PLS_INTEGER := NULL) IS
    t_sheet PLS_INTEGER := NVL(p_sheet
                              ,workbook.sheets.count());
  BEGIN
    workbook.sheets(t_sheet).freeze_rows := NULL;
    workbook.sheets(t_sheet).freeze_cols := p_nr_cols;
  END;
  --
  PROCEDURE freeze_pane(p_col   PLS_INTEGER
                       ,p_row   PLS_INTEGER
                       ,p_sheet PLS_INTEGER := NULL) IS
    t_sheet PLS_INTEGER := NVL(p_sheet
                              ,workbook.sheets.count());
  BEGIN
    workbook.sheets(t_sheet).freeze_rows := p_row;
    workbook.sheets(t_sheet).freeze_cols := p_col;
  END;
  --
  PROCEDURE set_autofilter(p_column_start PLS_INTEGER := NULL
                          ,p_column_end   PLS_INTEGER := NULL
                          ,p_row_start    PLS_INTEGER := NULL
                          ,p_row_end      PLS_INTEGER := NULL
                          ,p_sheet        PLS_INTEGER := NULL) IS
    t_ind   PLS_INTEGER;
    t_sheet PLS_INTEGER := NVL(p_sheet
                              ,workbook.sheets.count());
  BEGIN
    t_ind := 1;
    workbook.sheets(t_sheet).autofilters(t_ind).column_start := p_column_start;
    workbook.sheets(t_sheet).autofilters(t_ind).column_end := p_column_end;
    workbook.sheets(t_sheet).autofilters(t_ind).row_start := p_row_start;
    workbook.sheets(t_sheet).autofilters(t_ind).row_end := p_row_end;
    defined_name(p_column_start
                ,p_row_start
                ,p_column_end
                ,p_row_end
                ,'_xlnm._FilterDatabase'
                ,t_sheet
                ,t_sheet - 1);
  END;
  --
  /*
    procedure add1xml
      ( p_excel in out nocopy blob
      , p_filename varchar2
      , p_xml clob
      )
    is
      t_tmp blob;
      c_step constant number := 24396;
    begin
      dbms_lob.createtemporary( t_tmp, true );
      for i in 0 .. trunc( length( p_xml ) / c_step )
      loop
        dbms_lob.append( t_tmp, utl_i18n.string_to_raw( substr( p_xml, i * c_step + 1, c_step ), 'AL32UTF8' ) );
      end loop;
      add1file( p_excel, p_filename, t_tmp );
      dbms_lob.freetemporary( t_tmp );
    end;
  */
  --
  PROCEDURE add1xml(p_excel    IN OUT NOCOPY BLOB
                   ,p_filename VARCHAR2
                   ,p_xml      CLOB) IS
    t_tmp        BLOB;
    dest_offset  INTEGER := 1;
    src_offset   INTEGER := 1;
    lang_context INTEGER;
    warning      INTEGER;
  BEGIN
    lang_context := dbms_lob.DEFAULT_LANG_CTX;
    dbms_lob.createtemporary(t_tmp
                            ,TRUE);
    dbms_lob.converttoblob(t_tmp
                          ,p_xml
                          ,dbms_lob.lobmaxsize
                          ,dest_offset
                          ,src_offset
                          ,nls_charset_id('AL32UTF8') --
                          ,lang_context
                          ,warning);
    add1file(p_excel
            ,p_filename
            ,t_tmp);
    dbms_lob.freetemporary(t_tmp);
  END;
  --
  FUNCTION finish(p_landscape BOOLEAN DEFAULT FALSE) RETURN BLOB IS
    t_excel   BLOB;
    t_xxx     CLOB;
    t_tmp     VARCHAR2(32767 CHAR);
    t_str     VARCHAR2(32767 CHAR);
    t_c       NUMBER;
    t_h       NUMBER;
    t_w       NUMBER;
    t_cw      NUMBER;
    t_cell    VARCHAR2(1000 CHAR);
    t_row_ind PLS_INTEGER;
    t_col_min PLS_INTEGER;
    t_col_max PLS_INTEGER;
    t_col_ind PLS_INTEGER;
    t_len     PLS_INTEGER;
    ts        TIMESTAMP := systimestamp;
  BEGIN
    dbms_lob.createtemporary(t_excel
                            ,TRUE);
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
    FOR s IN 1 .. workbook.sheets.count()
    LOOP
      t_xxx := t_xxx || '
<Override PartName="/xl/worksheets/sheet' || s ||
               '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    END LOOP;
    t_xxx := t_xxx || '
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
    FOR s IN 1 .. workbook.sheets.count()
    LOOP
      IF workbook.sheets(s).comments.count() > 0
      THEN
        t_xxx := t_xxx || '
<Override PartName="/xl/comments' || s ||
                 '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>';
      END IF;
    END LOOP;
    t_xxx := t_xxx || '
</Types>';
    add1xml(t_excel
           ,'[Content_Types].xml'
           ,t_xxx);
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>' || sys_context('userenv'
                                        ,'os_user') || '</dc:creator>
<cp:lastModifiedBy>' ||
             sys_context('userenv'
                        ,'os_user') ||
             '</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">' ||
             to_char(current_timestamp
                    ,'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM') ||
             '</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">' ||
             to_char(current_timestamp
                    ,'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM') ||
             '</dcterms:modified>
</cp:coreProperties>';
    add1xml(t_excel
           ,'docProps/core.xml'
           ,t_xxx);
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Microsoft Excel</Application>
<DocSecurity>0</DocSecurity>
<ScaleCrop>false</ScaleCrop>
<HeadingPairs>
<vt:vector size="2" baseType="variant">
<vt:variant>
<vt:lpstr>Worksheets</vt:lpstr>
</vt:variant>
<vt:variant>
<vt:i4>' || workbook.sheets.count() || '</vt:i4>
</vt:variant>
</vt:vector>
</HeadingPairs>
<TitlesOfParts>
<vt:vector size="' || workbook.sheets.count() ||
             '" baseType="lpstr">';
    FOR s IN 1 .. workbook.sheets.count()
    LOOP
      t_xxx := t_xxx || '
<vt:lpstr>' || workbook.sheets(s).name || '</vt:lpstr>';
    END LOOP;
    t_xxx := t_xxx || '</vt:vector>
</TitlesOfParts>
<LinksUpToDate>false</LinksUpToDate>
<SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged>
<AppVersion>14.0300</AppVersion>
</Properties>';
    add1xml(t_excel
           ,'docProps/app.xml'
           ,t_xxx);
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>';
    add1xml(t_excel
           ,'_rels/.rels'
           ,t_xxx);
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">';
    IF workbook.numFmts.count() > 0
    THEN
      t_xxx := t_xxx || '<numFmts count="' || workbook.numFmts.count() || '">';
      FOR n IN 1 .. workbook.numFmts.count()
      LOOP
        t_xxx := t_xxx || '<numFmt numFmtId="' || workbook.numFmts(n)
                .numFmtId || '" formatCode="' || workbook.numFmts(n)
                .formatCode || '"/>';
      END LOOP;
      t_xxx := t_xxx || '</numFmts>';
    END IF;
    t_xxx := t_xxx || '<fonts count="' || workbook.fonts.count() ||
             '" x14ac:knownFonts="1">';
    FOR f IN 0 .. workbook.fonts.count() - 1
    LOOP
      t_xxx := t_xxx || '<font>' || CASE
                 WHEN workbook.fonts(f).bold THEN
                  '<b/>'
               END || CASE
                 WHEN workbook.fonts(f).italic THEN
                  '<i/>'
               END || CASE
                 WHEN workbook.fonts(f).underline THEN
                  '<u/>'
               END || '<sz val="' ||
               to_char(workbook.fonts(f).fontsize
                      ,'TM9'
                      ,'NLS_NUMERIC_CHARACTERS=.,') || '"/>
<color ' || CASE
                 WHEN workbook.fonts(f).rgb IS NOT NULL THEN
                  'rgb="' || workbook.fonts(f).rgb
                 ELSE
                  'theme="' || workbook.fonts(f).theme
               END || '"/>
<name val="' || workbook.fonts(f).name || '"/>
<family val="' || workbook.fonts(f).family || '"/>
<scheme val="none"/>
</font>';
    END LOOP;
    t_xxx := t_xxx || '</fonts>
<fills count="' || workbook.fills.count() || '">';
    FOR f IN 0 .. workbook.fills.count() - 1
    LOOP
      t_xxx := t_xxx || '<fill><patternFill patternType="' || workbook.fills(f)
              .patternType || '">' || CASE
                 WHEN workbook.fills(f).fgRGB IS NOT NULL THEN
                  '<fgColor rgb="' || workbook.fills(f).fgRGB || '"/>'
               END || '</patternFill></fill>';
    END LOOP;
    t_xxx := t_xxx || '</fills>
<borders count="' || workbook.borders.count() || '">';
    FOR b IN 0 .. workbook.borders.count() - 1
    LOOP
      t_xxx := t_xxx || '<border>' || CASE
                 WHEN workbook.borders(b).left IS NULL THEN
                  '<left/>'
                 ELSE
                  '<left style="' || workbook.borders(b).left || '"/>'
               END || CASE
                 WHEN workbook.borders(b).right IS NULL THEN
                  '<right/>'
                 ELSE
                  '<right style="' || workbook.borders(b).right || '"/>'
               END || CASE
                 WHEN workbook.borders(b).top IS NULL THEN
                  '<top/>'
                 ELSE
                  '<top style="' || workbook.borders(b).top || '"/>'
               END || CASE
                 WHEN workbook.borders(b).bottom IS NULL THEN
                  '<bottom/>'
                 ELSE
                  '<bottom style="' || workbook.borders(b).bottom || '"/>'
               END || '</border>';
    END LOOP;
    t_xxx := t_xxx || '</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="' || (workbook.cellXfs.count() + 1) || '">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    FOR x IN 1 .. workbook.cellXfs.count()
    LOOP
      t_xxx := t_xxx || '<xf numFmtId="' || workbook.cellXfs(x).numFmtId ||
               '" fontId="' || workbook.cellXfs(x).fontId || '" fillId="' || workbook.cellXfs(x)
              .fillId || '" borderId="' || workbook.cellXfs(x).borderId || '">';
      IF (workbook.cellXfs(x).alignment.horizontal IS NOT NULL OR workbook.cellXfs(x)
         .alignment.vertical IS NOT NULL OR workbook.cellXfs(x)
         .alignment.wrapText)
      THEN
        t_xxx := t_xxx || '<alignment' || CASE
                   WHEN workbook.cellXfs(x).alignment.horizontal IS NOT NULL THEN
                    ' horizontal="' || workbook.cellXfs(x).alignment.horizontal || '"'
                 END || CASE
                   WHEN workbook.cellXfs(x).alignment.vertical IS NOT NULL THEN
                    ' vertical="' || workbook.cellXfs(x).alignment.vertical || '"'
                 END || CASE
                   WHEN workbook.cellXfs(x).alignment.wrapText THEN
                    ' wrapText="true"'
                 END || '/>';
      END IF;
      t_xxx := t_xxx || '</xf>';
    END LOOP;
    t_xxx := t_xxx || '</cellXfs>
<cellStyles count="1">
<cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
<dxfs count="0"/>
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
<extLst>
<ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
</ext>
</extLst>
</styleSheet>';
    add1xml(t_excel
           ,'xl/styles.xml'
           ,t_xxx);
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/>
<workbookPr date1904="true" defaultThemeVersion="124226"/>
<bookViews>
<workbookView xWindow="120" yWindow="45" windowWidth="19155" windowHeight="4935"/>
</bookViews>
<sheets>';
    FOR s IN 1 .. workbook.sheets.count()
    LOOP
      t_xxx := t_xxx || '
<sheet name="' || workbook.sheets(s).name || '" sheetId="' || s ||
               '" r:id="rId' || (9 + s) || '"/>';
    END LOOP;
    t_xxx := t_xxx || '</sheets>';
    IF workbook.defined_names.count() > 0
    THEN
      t_xxx := t_xxx || '<definedNames>';
      FOR s IN 1 .. workbook.defined_names.count()
      LOOP
        t_xxx := t_xxx || '
<definedName name="' || workbook.defined_names(s).name || '"' || CASE
                   WHEN workbook.defined_names(s).sheet IS NOT NULL THEN
                    ' localSheetId="' || to_char(workbook.defined_names(s).sheet) || '"'
                 END || '>' || workbook.defined_names(s).ref ||
                 '</definedName>';
      END LOOP;
      t_xxx := t_xxx || '</definedNames>';
    END IF;
    t_xxx := t_xxx || '<calcPr calcId="144525"/></workbook>';
    add1xml(t_excel
           ,'xl/workbook.xml'
           ,t_xxx);
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
<a:themeElements>
<a:clrScheme name="Office">
<a:dk1>
<a:sysClr val="windowText" lastClr="000000"/>
</a:dk1>
<a:lt1>
<a:sysClr val="window" lastClr="FFFFFF"/>
</a:lt1>
<a:dk2>
<a:srgbClr val="1F497D"/>
</a:dk2>
<a:lt2>
<a:srgbClr val="EEECE1"/>
</a:lt2>
<a:accent1>
<a:srgbClr val="4F81BD"/>
</a:accent1>
<a:accent2>
<a:srgbClr val="C0504D"/>
</a:accent2>
<a:accent3>
<a:srgbClr val="9BBB59"/>
</a:accent3>
<a:accent4>
<a:srgbClr val="8064A2"/>
</a:accent4>
<a:accent5>
<a:srgbClr val="4BACC6"/>
</a:accent5>
<a:accent6>
<a:srgbClr val="F79646"/>
</a:accent6>
<a:hlink>
<a:srgbClr val="0000FF"/>
</a:hlink>
<a:folHlink>
<a:srgbClr val="800080"/>
</a:folHlink>
</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont>
<a:latin typeface="Cambria"/>
<a:ea typeface=""/>
<a:cs typeface=""/>
<a:font script="Jpan" typeface="MS P????"/>
<a:font script="Hang" typeface="?? ??"/>
<a:font script="Hans" typeface="??"/>
<a:font script="Hant" typeface="????"/>
<a:font script="Arab" typeface="Times New Roman"/>
<a:font script="Hebr" typeface="Times New Roman"/>
<a:font script="Thai" typeface="Tahoma"/>
<a:font script="Ethi" typeface="Nyala"/>
<a:font script="Beng" typeface="Vrinda"/>
<a:font script="Gujr" typeface="Shruti"/>
<a:font script="Khmr" typeface="MoolBoran"/>
<a:font script="Knda" typeface="Tunga"/>
<a:font script="Guru" typeface="Raavi"/>
<a:font script="Cans" typeface="Euphemia"/>
<a:font script="Cher" typeface="Plantagenet Cherokee"/>
<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
<a:font script="Tibt" typeface="Microsoft Himalaya"/>
<a:font script="Thaa" typeface="MV Boli"/>
<a:font script="Deva" typeface="Mangal"/>
<a:font script="Telu" typeface="Gautami"/>
<a:font script="Taml" typeface="Latha"/>
<a:font script="Syrc" typeface="Estrangelo Edessa"/>
<a:font script="Orya" typeface="Kalinga"/>
<a:font script="Mlym" typeface="Kartika"/>
<a:font script="Laoo" typeface="DokChampa"/>
<a:font script="Sinh" typeface="Iskoola Pota"/>
<a:font script="Mong" typeface="Mongolian Baiti"/>
<a:font script="Viet" typeface="Times New Roman"/>
<a:font script="Uigh" typeface="Microsoft Uighur"/>
<a:font script="Geor" typeface="Sylfaen"/>
</a:majorFont>
<a:minorFont>
<a:latin typeface="Calibri"/>
<a:ea typeface=""/>
<a:cs typeface=""/>
<a:font script="Jpan" typeface="MS P????"/>
<a:font script="Hang" typeface="?? ??"/>
<a:font script="Hans" typeface="??"/>
<a:font script="Hant" typeface="????"/>
<a:font script="Arab" typeface="Arial"/>
<a:font script="Hebr" typeface="Arial"/>
<a:font script="Thai" typeface="Tahoma"/>
<a:font script="Ethi" typeface="Nyala"/>
<a:font script="Beng" typeface="Vrinda"/>
<a:font script="Gujr" typeface="Shruti"/>
<a:font script="Khmr" typeface="DaunPenh"/>
<a:font script="Knda" typeface="Tunga"/>
<a:font script="Guru" typeface="Raavi"/>
<a:font script="Cans" typeface="Euphemia"/>
<a:font script="Cher" typeface="Plantagenet Cherokee"/>
<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
<a:font script="Tibt" typeface="Microsoft Himalaya"/>
<a:font script="Thaa" typeface="MV Boli"/>
<a:font script="Deva" typeface="Mangal"/>
<a:font script="Telu" typeface="Gautami"/>
<a:font script="Taml" typeface="Latha"/>
<a:font script="Syrc" typeface="Estrangelo Edessa"/>
<a:font script="Orya" typeface="Kalinga"/>
<a:font script="Mlym" typeface="Kartika"/>
<a:font script="Laoo" typeface="DokChampa"/>
<a:font script="Sinh" typeface="Iskoola Pota"/>
<a:font script="Mong" typeface="Mongolian Baiti"/>
<a:font script="Viet" typeface="Arial"/>
<a:font script="Uigh" typeface="Microsoft Uighur"/>
<a:font script="Geor" typeface="Sylfaen"/>
</a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst>
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="50000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="35000">
<a:schemeClr val="phClr">
<a:tint val="37000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:tint val="15000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:lin ang="16200000" scaled="1"/>
</a:gradFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:shade val="51000"/>
<a:satMod val="130000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="80000">
<a:schemeClr val="phClr">
<a:shade val="93000"/>
<a:satMod val="130000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="94000"/>
<a:satMod val="135000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:lin ang="16200000" scaled="0"/>
</a:gradFill>
</a:fillStyleLst>
<a:lnStyleLst>
<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr">
<a:shade val="95000"/>
<a:satMod val="105000"/>
</a:schemeClr>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
</a:lnStyleLst>
<a:effectStyleLst>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="38000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
</a:effectStyle>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="35000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
</a:effectStyle>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="35000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
<a:scene3d>
<a:camera prst="orthographicFront">
<a:rot lat="0" lon="0" rev="0"/>
</a:camera>
<a:lightRig rig="threePt" dir="t">
<a:rot lat="0" lon="0" rev="1200000"/>
</a:lightRig>
</a:scene3d>
<a:sp3d>
<a:bevelT w="63500" h="25400"/>
</a:sp3d>
</a:effectStyle>
</a:effectStyleLst>
<a:bgFillStyleLst>
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="40000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="40000">
<a:schemeClr val="phClr">
<a:tint val="45000"/>
<a:shade val="99000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="20000"/>
<a:satMod val="255000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:path path="circle">
<a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
</a:path>
</a:gradFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="80000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="30000"/>
<a:satMod val="200000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:path path="circle">
<a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
</a:path>
</a:gradFill>
</a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
<a:objectDefaults/>
<a:extraClrSchemeLst/>
</a:theme>';
    add1xml(t_excel
           ,'xl/theme/theme1.xml'
           ,t_xxx);
    FOR s IN 1 .. workbook.sheets.count()
    LOOP
      t_col_min := 16384;
      t_col_max := 1;
      t_row_ind := workbook.sheets(s).rows.first();
      WHILE t_row_ind IS NOT NULL
      LOOP
        t_col_min := least(t_col_min
                          ,workbook.sheets(s).rows(t_row_ind).first());
        t_col_max := greatest(t_col_max
                             ,workbook.sheets(s).rows(t_row_ind).last());
        t_row_ind := workbook.sheets(s).rows.next(t_row_ind);
      END LOOP;
      t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
<dimension ref="' || alfan_col(t_col_min) || workbook.sheets(s)
              .rows.first() || ':' || alfan_col(t_col_max) || workbook.sheets(s)
              .rows.last() || '"/>
<sheetViews>
<sheetView' || CASE
                 WHEN s = 1 THEN
                  ' tabSelected="1"'
               END || ' workbookViewId="0">';
      IF workbook.sheets(s).freeze_rows > 0
          AND workbook.sheets(s).freeze_cols > 0
      THEN
        t_xxx := t_xxx || ('<pane xSplit="' || workbook.sheets(s).freeze_cols || '" ' ||
                 'ySplit="' || workbook.sheets(s).freeze_rows || '" ' ||
                 'topLeftCell="' ||
                 alfan_col(workbook.sheets(s).freeze_cols + 1) ||
                 (workbook.sheets(s).freeze_rows + 1) || '" ' ||
                 'activePane="bottomLeft" state="frozen"/>');
      ELSE
        IF workbook.sheets(s).freeze_rows > 0
        THEN
          t_xxx := t_xxx || '<pane ySplit="' || workbook.sheets(s)
                  .freeze_rows || '" topLeftCell="A' ||
                   (workbook.sheets(s).freeze_rows + 1) ||
                   '" activePane="bottomLeft" state="frozen"/>';
        END IF;
        IF workbook.sheets(s).freeze_cols > 0
        THEN
          t_xxx := t_xxx || '<pane xSplit="' || workbook.sheets(s)
                  .freeze_cols || '" topLeftCell="' ||
                   alfan_col(workbook.sheets(s).freeze_cols + 1) ||
                   '1" activePane="bottomLeft" state="frozen"/>';
        END IF;
      END IF;
      t_xxx := t_xxx ||
               '</sheetView>
</sheetViews>
<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>';
      IF workbook.sheets(s).widths.count() > 0
      THEN
        t_xxx     := t_xxx || '<cols>';
        t_col_ind := workbook.sheets(s).widths.first();
        WHILE t_col_ind IS NOT NULL
        LOOP
          t_xxx     := t_xxx || '<col min="' || t_col_ind || '" max="' ||
                       t_col_ind || '" width="' ||
                       to_char(workbook.sheets(s).widths(t_col_ind)
                              ,'TM9'
                              ,'NLS_NUMERIC_CHARACTERS=.,') ||
                       '" customWidth="1"/>';
          t_col_ind := workbook.sheets(s).widths.next(t_col_ind);
        END LOOP;
        t_xxx := t_xxx || '</cols>';
      END IF;
      t_xxx     := t_xxx || '<sheetData>';
      t_row_ind := workbook.sheets(s).rows.first();
      t_tmp     := NULL;
      WHILE t_row_ind IS NOT NULL
      LOOP
        t_tmp     := t_tmp || '<row r="' || t_row_ind || '" spans="' ||
                     t_col_min || ':' || t_col_max || '">';
        t_len     := length(t_tmp);
        t_col_ind := workbook.sheets(s).rows(t_row_ind).first();
        WHILE t_col_ind IS NOT NULL
        LOOP
          t_cell := '<c r="' || alfan_col(t_col_ind) || t_row_ind || '"' || ' ' || workbook.sheets(s)
                   .rows(t_row_ind)(t_col_ind)
                   .style || '><v>' ||
                    to_char(workbook.sheets(s).rows(t_row_ind)(t_col_ind)
                            .value
                           ,'TM9'
                           ,'NLS_NUMERIC_CHARACTERS=.,') || '</v></c>';
          IF t_len > 32000
          THEN
            dbms_lob.writeappend(t_xxx
                                ,t_len
                                ,t_tmp);
            t_tmp := NULL;
            t_len := 0;
          END IF;
          t_tmp     := t_tmp || t_cell;
          t_len     := t_len + length(t_cell);
          t_col_ind := workbook.sheets(s).rows(t_row_ind).next(t_col_ind);
        END LOOP;
        t_tmp     := t_tmp || '</row>';
        t_row_ind := workbook.sheets(s).rows.next(t_row_ind);
      END LOOP;
      t_tmp := t_tmp || '</sheetData>';
      t_len := length(t_tmp);
      dbms_lob.writeappend(t_xxx
                          ,t_len
                          ,t_tmp);
      FOR a IN 1 .. workbook.sheets(s).autofilters.count()
      LOOP
        t_xxx := t_xxx || '<autoFilter ref="' ||
                 alfan_col(NVL(workbook.sheets(s).autofilters(a).column_start
                              ,t_col_min)) ||
                 NVL(workbook.sheets(s).autofilters(a).row_start
                    ,workbook.sheets(s).rows.first()) || ':' ||
                 alfan_col(coalesce(workbook.sheets(s).autofilters(a)
                                    .column_end
                                   ,workbook.sheets(s).autofilters(a)
                                    .column_start
                                   ,t_col_max)) ||
                 NVL(workbook.sheets(s).autofilters(a).row_end
                    ,workbook.sheets(s).rows.last()) || '"/>';
      END LOOP;
      IF workbook.sheets(s).mergecells.count() > 0
      THEN
        t_xxx := t_xxx || '<mergeCells count="' ||
                 to_char(workbook.sheets(s).mergecells.count()) || '">';
        FOR m IN 1 .. workbook.sheets(s).mergecells.count()
        LOOP
          t_xxx := t_xxx || '<mergeCell ref="' || workbook.sheets(s)
                  .mergecells(m) || '"/>';
        END LOOP;
        t_xxx := t_xxx || '</mergeCells>';
      END IF;
      --
      IF workbook.sheets(s).validations.count() > 0
      THEN
        t_xxx := t_xxx || '<dataValidations count="' ||
                 to_char(workbook.sheets(s).validations.count()) || '">';
        FOR m IN 1 .. workbook.sheets(s).validations.count()
        LOOP
          t_xxx := t_xxx || '<dataValidation' || ' type="' || workbook.sheets(s).validations(m).type || '"' ||
                   ' errorStyle="' || workbook.sheets(s).validations(m)
                  .errorstyle || '"' || ' allowBlank="' || CASE
                     WHEN NVL(workbook.sheets(s).validations(m).allowBlank
                             ,TRUE) THEN
                      '1'
                     ELSE
                      '0'
                   END || '"' || ' sqref="' || workbook.sheets(s).validations(m)
                  .sqref || '"';
          IF workbook.sheets(s).validations(m).prompt IS NOT NULL
          THEN
            t_xxx := t_xxx || ' showInputMessage="1" prompt="' || workbook.sheets(s).validations(m)
                    .prompt || '"';
            IF workbook.sheets(s).validations(m).title IS NOT NULL
            THEN
              t_xxx := t_xxx || ' promptTitle="' || workbook.sheets(s).validations(m)
                      .title || '"';
            END IF;
          END IF;
          IF workbook.sheets(s).validations(m).showerrormessage
          THEN
            t_xxx := t_xxx || ' showErrorMessage="1"';
            IF workbook.sheets(s).validations(m).error_title IS NOT NULL
            THEN
              t_xxx := t_xxx || ' errorTitle="' || workbook.sheets(s).validations(m)
                      .error_title || '"';
            END IF;
            IF workbook.sheets(s).validations(m).error_txt IS NOT NULL
            THEN
              t_xxx := t_xxx || ' error="' || workbook.sheets(s).validations(m)
                      .error_txt || '"';
            END IF;
          END IF;
          t_xxx := t_xxx || '>';
          IF workbook.sheets(s).validations(m).formula1 IS NOT NULL
          THEN
            t_xxx := t_xxx || '<formula1>' || workbook.sheets(s).validations(m)
                    .formula1 || '</formula1>';
          END IF;
          IF workbook.sheets(s).validations(m).formula2 IS NOT NULL
          THEN
            t_xxx := t_xxx || '<formula2>' || workbook.sheets(s).validations(m)
                    .formula2 || '</formula2>';
          END IF;
          t_xxx := t_xxx || '</dataValidation>';
        END LOOP;
        t_xxx := t_xxx || '</dataValidations>';
      END IF;
      --
      IF workbook.sheets(s).hyperlinks.count() > 0
      THEN
        t_xxx := t_xxx || '<hyperlinks>';
        FOR h IN 1 .. workbook.sheets(s).hyperlinks.count()
        LOOP
          t_xxx := t_xxx || '<hyperlink ref="' || workbook.sheets(s).hyperlinks(h).cell ||
                   '" r:id="rId' || h || '"/>';
        END LOOP;
        t_xxx := t_xxx || '</hyperlinks>';
      END IF;
      t_xxx := t_xxx ||
               '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';
      -- portrait / landscape
      IF p_landscape
      THEN
        t_xxx := t_xxx ||
                 '<pageSetup paperSize="9" orientation="landscape"/>';
      ELSE
        t_xxx := t_xxx || '<pageSetup paperSize="9" orientation="portrait"/>';
      END IF;
      --
      IF workbook.sheets(s).comments.count() > 0
      THEN
        t_xxx := t_xxx || '<legacyDrawing r:id="rId' ||
                 (workbook.sheets(s).hyperlinks.count() + 1) || '"/>';
      END IF;
      --
      t_xxx := t_xxx || '</worksheet>';
      add1xml(t_excel
             ,'xl/worksheets/sheet' || s || '.xml'
             ,t_xxx);
      IF workbook.sheets(s).hyperlinks.count() > 0
          OR workbook.sheets(s).comments.count() > 0
      THEN
        t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        IF workbook.sheets(s).comments.count() > 0
        THEN
          t_xxx := t_xxx || '<Relationship Id="rId' ||
                   (workbook.sheets(s).hyperlinks.count() + 2) ||
                   '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments' || s ||
                   '.xml"/>';
          t_xxx := t_xxx || '<Relationship Id="rId' ||
                   (workbook.sheets(s).hyperlinks.count() + 1) ||
                   '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing' || s ||
                   '.vml"/>';
        END IF;
        FOR h IN 1 .. workbook.sheets(s).hyperlinks.count()
        LOOP
          t_xxx := t_xxx || '<Relationship Id="rId' || h ||
                   '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' || workbook.sheets(s).hyperlinks(h).url ||
                   '" TargetMode="External"/>';
        END LOOP;
        t_xxx := t_xxx || '</Relationships>';
        add1xml(t_excel
               ,'xl/worksheets/_rels/sheet' || s || '.xml.rels'
               ,t_xxx);
      END IF;
      --
      IF workbook.sheets(s).comments.count() > 0
      THEN
        DECLARE
          cnt        PLS_INTEGER;
          author_ind tp_author;
          --          t_col_ind := workbook.sheets( s ).widths.next( t_col_ind );
        BEGIN
          authors.delete();
          FOR c IN 1 .. workbook.sheets(s).comments.count()
          LOOP
            authors(workbook.sheets(s).comments(c).author) := 0;
          END LOOP;
          t_xxx      := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors>';
          cnt        := 0;
          author_ind := authors.first();
          WHILE author_ind IS NOT NULL
                OR authors.next(author_ind) IS NOT NULL
          LOOP
            authors(author_ind) := cnt;
            t_xxx := t_xxx || '<author>' || author_ind || '</author>';
            cnt := cnt + 1;
            author_ind := authors.next(author_ind);
          END LOOP;
        END;
        t_xxx := t_xxx || '</authors><commentList>';
        FOR c IN 1 .. workbook.sheets(s).comments.count()
        LOOP
          t_xxx := t_xxx || '<comment ref="' ||
                   alfan_col(workbook.sheets(s).comments(c).column) ||
                   to_char(workbook.sheets(s).comments(c)
                           .row || '" authorId="' ||
                            authors(workbook.sheets(s).comments(c).author)) || '">
<text>';
          IF workbook.sheets(s).comments(c).author IS NOT NULL
          THEN
            t_xxx := t_xxx ||
                     '<r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t xml:space="preserve">' || workbook.sheets(s).comments(c)
                    .author || ':</t></r>';
          END IF;
          t_xxx := t_xxx ||
                   '<r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t xml:space="preserve">' || CASE
                     WHEN workbook.sheets(s).comments(c).author IS NOT NULL THEN
                      '
'
                   END || workbook.sheets(s).comments(c).text ||
                   '</t></r></text></comment>';
        END LOOP;
        t_xxx := t_xxx || '</commentList></comments>';
        add1xml(t_excel
               ,'xl/comments' || s || '.xml'
               ,t_xxx);
        t_xxx := '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="2"/></o:shapelayout>
<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';
        FOR c IN 1 .. workbook.sheets(s).comments.count()
        LOOP
          t_xxx := t_xxx || '<v:shape id="_x0000_s' || to_char(c) ||
                   '" type="#_x0000_t202"
style="position:absolute;margin-left:35.25pt;margin-top:3pt;z-index:' ||
                   to_char(c) ||
                   ';visibility:hidden;" fillcolor="#ffffe1" o:insetmode="auto">
<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>
<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>
<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>';
          t_w   := workbook.sheets(s).comments(c).width;
          t_c   := 1;
          LOOP
            IF workbook.sheets(s)
             .widths.exists(workbook.sheets(s).comments(c).column + t_c)
            THEN
              t_cw := 256 * workbook.sheets(s)
                     .widths(workbook.sheets(s).comments(c).column + t_c);
              t_cw := trunc((t_cw + 18) / 256 * 7); -- assume default 11 point Calibri
            ELSE
              t_cw := 64;
            END IF;
            EXIT WHEN t_w < t_cw;
            t_c := t_c + 1;
            t_w := t_w - t_cw;
          END LOOP;
          t_h   := workbook.sheets(s).comments(c).height;
          t_xxx := t_xxx || to_char('<x:Anchor>' || workbook.sheets(s).comments(c)
                                    .column || ',15,' || workbook.sheets(s).comments(c).row ||
                                    ',30,' || (workbook.sheets(s).comments(c)
                                    .column + t_c - 1) || ',' || round(t_w) || ',' ||
                                    (workbook.sheets(s).comments(c)
                                    .row + 1 + trunc(t_h / 20)) || ',' ||
                                    MOD(t_h
                                       ,20) || '</x:Anchor>');
          t_xxx := t_xxx ||
                   to_char('<x:AutoFill>False</x:AutoFill><x:Row>' ||
                           (workbook.sheets(s).comments(c).row - 1) ||
                           '</x:Row><x:Column>' ||
                           (workbook.sheets(s).comments(c).column - 1) ||
                           '</x:Column></x:ClientData></v:shape>');
        END LOOP;
        t_xxx := t_xxx || '</xml>';
        add1xml(t_excel
               ,'xl/drawings/vmlDrawing' || s || '.vml'
               ,t_xxx);
      END IF;
      --
    END LOOP;
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>';
    FOR s IN 1 .. workbook.sheets.count()
    LOOP
      t_xxx := t_xxx || '
<Relationship Id="rId' || (9 + s) ||
               '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' || s ||
               '.xml"/>';
    END LOOP;
    t_xxx := t_xxx || '</Relationships>';
    add1xml(t_excel
           ,'xl/_rels/workbook.xml.rels'
           ,t_xxx);
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' ||
             workbook.str_cnt || '" uniqueCount="' ||
             workbook.strings.count() || '">';
    t_tmp := NULL;
    FOR i IN 0 .. workbook.str_ind.count() - 1
    LOOP
      t_str := '<si><t>' || dbms_xmlgen.convert(substr(workbook.str_ind(i)
                                                      ,1
                                                      ,32000)) || '</t></si>';
      IF length(t_tmp) + length(t_str) > 32000
      THEN
        t_xxx := t_xxx || t_tmp;
        t_tmp := NULL;
      END IF;
      t_tmp := t_tmp || t_str;
    END LOOP;
    t_xxx := t_xxx || t_tmp || '</sst>';
    add1xml(t_excel
           ,'xl/sharedStrings.xml'
           ,t_xxx);
    finish_zip(t_excel);
    clear_workbook;
    RETURN t_excel;
  END;
  --
  PROCEDURE SAVE(p_directory VARCHAR2
                ,p_filename  VARCHAR2) IS
  BEGIN
    blob2file(finish
             ,p_directory
             ,p_filename);
  END;
  --
  PROCEDURE query2sheet(p_sql            VARCHAR2
                       ,p_column_headers BOOLEAN := TRUE
                       ,p_directory      VARCHAR2 := NULL
                       ,p_filename       VARCHAR2 := NULL
                       ,p_sheet          PLS_INTEGER := NULL) IS
    t_sheet     PLS_INTEGER;
    t_c         INTEGER;
    t_col_cnt   INTEGER;
    t_desc_tab  dbms_sql.desc_tab2;
    d_tab       dbms_sql.date_table;
    n_tab       dbms_sql.number_table;
    v_tab       dbms_sql.varchar2_table;
    t_bulk_size PLS_INTEGER := 200;
    t_r         INTEGER;
    t_cur_row   PLS_INTEGER;
  BEGIN
    IF p_sheet IS NULL
    THEN
      new_sheet;
    END IF;
    t_c := dbms_sql.open_cursor;
    dbms_sql.parse(t_c
                  ,p_sql
                  ,dbms_sql.native);
    dbms_sql.describe_columns2(t_c
                              ,t_col_cnt
                              ,t_desc_tab);
    FOR c IN 1 .. t_col_cnt
    LOOP
      IF p_column_headers
      THEN
        cell(c
            ,1
            ,t_desc_tab(c).col_name
            ,p_sheet    => t_sheet);
      END IF;
      --      dbms_output.put_line( t_desc_tab( c ).col_name || ' ' || t_desc_tab( c ).col_type );
      CASE
        WHEN t_desc_tab(c).col_type IN (2
                          ,100
                          ,101) THEN
          dbms_sql.define_array(t_c
                               ,c
                               ,n_tab
                               ,t_bulk_size
                               ,1);
        WHEN t_desc_tab(c).col_type IN (12
                          ,178
                          ,179
                          ,180
                          ,181
                          ,231) THEN
          dbms_sql.define_array(t_c
                               ,c
                               ,d_tab
                               ,t_bulk_size
                               ,1);
        WHEN t_desc_tab(c).col_type IN (1
                          ,8
                          ,9
                          ,96
                          ,112) THEN
          dbms_sql.define_array(t_c
                               ,c
                               ,v_tab
                               ,t_bulk_size
                               ,1);
        ELSE
          NULL;
      END CASE;
    END LOOP;
    --
    t_cur_row := CASE
                   WHEN p_column_headers THEN
                    2
                   ELSE
                    1
                 END;
    t_sheet   := NVL(p_sheet
                    ,workbook.sheets.count());
    --
    t_r := dbms_sql.execute(t_c);
    LOOP
      t_r := dbms_sql.fetch_rows(t_c);
      IF t_r > 0
      THEN
        FOR c IN 1 .. t_col_cnt
        LOOP
          CASE
            WHEN t_desc_tab(c).col_type IN (2
                              ,100
                              ,101) THEN
              dbms_sql.column_value(t_c
                                   ,c
                                   ,n_tab);
              FOR i IN 0 .. t_r - 1
              LOOP
                IF n_tab(i + n_tab.first()) IS NOT NULL
                THEN
                  cell(c
                      ,t_cur_row + i
                      ,n_tab(i + n_tab.first())
                      ,p_sheet => t_sheet);
                END IF;
              END LOOP;
              n_tab.delete;
            WHEN t_desc_tab(c).col_type IN (12
                              ,178
                              ,179
                              ,180
                              ,181
                              ,231) THEN
              dbms_sql.column_value(t_c
                                   ,c
                                   ,d_tab);
              FOR i IN 0 .. t_r - 1
              LOOP
                IF d_tab(i + d_tab.first()) IS NOT NULL
                THEN
                  cell(c
                      ,t_cur_row + i
                      ,d_tab(i + d_tab.first())
                      ,p_sheet => t_sheet);
                END IF;
              END LOOP;
              d_tab.delete;
            WHEN t_desc_tab(c).col_type IN (1
                              ,8
                              ,9
                              ,96
                              ,112) THEN
              dbms_sql.column_value(t_c
                                   ,c
                                   ,v_tab);
              FOR i IN 0 .. t_r - 1
              LOOP
                IF v_tab(i + v_tab.first()) IS NOT NULL
                THEN
                  cell(c
                      ,t_cur_row + i
                      ,v_tab(i + v_tab.first())
                      ,p_sheet => t_sheet);
                END IF;
              END LOOP;
              v_tab.delete;
            ELSE
              NULL;
          END CASE;
        END LOOP;
      END IF;
      EXIT WHEN t_r != t_bulk_size;
      t_cur_row := t_cur_row + t_r;
    END LOOP;
    dbms_sql.close_cursor(t_c);
    IF (p_directory IS NOT NULL AND p_filename IS NOT NULL)
    THEN
      SAVE(p_directory
          ,p_filename);
    END IF;
  EXCEPTION
    WHEN OTHERS THEN
      IF dbms_sql.is_open(t_c)
      THEN
        dbms_sql.close_cursor(t_c);
      END IF;
  END;
  --
  FUNCTION create_xlsx_apex(p_process IN apex_plugin.t_process
                           ,p_plugin  IN apex_plugin.t_plugin)
    RETURN apex_plugin.t_process_exec_result IS
    c_max_columns         PLS_INTEGER := 150;
    p_headers             BOOLEAN;
    p_sheetname           BOOLEAN;
    p_regionname          VARCHAR2(32767);
    p_filename            VARCHAR2(32767);
    p_user_filename       VARCHAR2(32767);
    p_append_sysdate      VARCHAR2(32767);
    p_exclude_column_list VARCHAR2(32767);
    p_use_filters         BOOLEAN;
    t_app_id              NUMBER;
    t_page_id             NUMBER;
    t_session_id          NUMBER;
    --
    l_sheet_name VARCHAR2(32767);
    l_excel      BLOB;
    l_sheet_id   PLS_INTEGER;
    l_filename   VARCHAR2(32767);
    l_order_by   VARCHAR2(32767);
    l_row        NUMBER;
    l_column     NUMBER;

    t_val apex_plugin_util.t_column_value_list2;
    TYPE tp_col IS RECORD(
       label apex_application_page_ir_col.report_label%TYPE
      ,alias apex_application_page_ir_col.column_alias%TYPE
      ,ind   PLS_INTEGER);
    TYPE tp_cols IS TABLE OF tp_col;
    t_cols tp_cols;
  BEGIN
    apex_plugin_util.debug_process(p_plugin
                                  ,p_process);
    --
    p_headers             := p_process.attribute_01 = 'Y';
    p_sheetname           := p_process.attribute_02 = 'Y';
    p_regionname          := p_process.attribute_03;
    p_filename            := p_process.attribute_04;
    p_user_filename       := p_process.attribute_05;
    p_append_sysdate      := p_process.attribute_06;
    p_exclude_column_list := p_process.attribute_07;
    p_use_filters         := p_process.attribute_08 = 'Y';
    t_app_id              := v('APP_ID');
    t_page_id             := v('APP_PAGE_ID');
    t_session_id          := v('APP_SESSION');
    --
    as_xlsx.clear_workbook;

    --
    apex_debug_message.log_message('looking for report regions');
    l_sheet_id := 1;
    FOR r_apr IN (SELECT apr.region_id
                        ,apr.region_name
                        ,apr.region_source
                        ,apr.source_type_code
                        ,apr.report_column_source_type
                        ,apr.page_name
                    FROM apex_application_page_regions apr
                   WHERE apr.application_id = t_app_id
                     AND apr.page_id = t_page_id
                     AND apr.source_type_code IN
                         ('SQL_QUERY'
                         ,'DYNAMIC_QUERY'
                         ,'STRUCTURED_QUERY'
                         ,'FUNCTION_RETURNING_SQL_QUERY')
                     AND apr.region_name IN
                         (SELECT substr(prn
                                       ,instr(';' || prn
                                             ,';'
                                             ,1
                                             ,LEVEL)
                                       ,instr(prn || ';'
                                             ,';'
                                             ,1
                                             ,LEVEL) -
                                        instr(';' || prn
                                             ,';'
                                             ,1
                                             ,LEVEL))
                            FROM (SELECT p_regionname prn
                                    FROM dual)
                          CONNECT BY LEVEL <=
                                     length(prn) - NVL(length(REPLACE(prn
                                                                     ,';'))
                                                      ,0) + 1
                          UNION ALL
                          SELECT apr.region_name
                            FROM dual
                           WHERE p_regionname IS NULL)
                   ORDER BY apr.display_sequence)
    LOOP
      --

      --
      apex_debug_message.log_message('found ' || r_apr.source_type_code || ' ' ||
                                     r_apr.region_name);
      --
      IF r_apr.source_type_code = 'FUNCTION_RETURNING_SQL_QUERY'
      THEN
        r_apr.region_source := apex_plugin_util.get_plsql_function_result(r_apr.region_source);
      END IF;
      --
      IF r_apr.source_type_code = 'DYNAMIC_QUERY'
      THEN
        IF p_use_filters
        THEN
          DECLARE
            t_base_report_id        NUMBER;
            t_interactive_report_id NUMBER;
            t_pref                  VARCHAR2(32767);
            t_where                 CLOB;
            t_val_ind               PLS_INTEGER;
            t_start                 PLS_INTEGER;
            t_end                   PLS_INTEGER;
            t_val                   VARCHAR2(32767);
            t_search                VARCHAR2(32767);
            t_apr                   apex_application_page_ir_rpt%ROWTYPE;
          BEGIN
            SELECT api.interactive_report_id
              INTO t_interactive_report_id
              FROM apex_application_page_ir api
             WHERE api.application_id = t_app_id
               AND api.page_id = t_page_id
               AND api.region_id = r_apr.region_id;
            --
            t_pref           := apex_util.get_preference('FSP_IR_' ||
                                                         t_app_id || '_P' ||
                                                         t_page_id || '_W' ||
                                                         t_interactive_report_id
                                                        ,v('APP_USER'));
            t_base_report_id := substr(t_pref
                                      ,1
                                      ,instr(t_pref
                                            ,'_') - 1);
            --
            SELECT *
              INTO t_apr
              FROM apex_application_page_ir_rpt apr
             WHERE apr.application_id = t_app_id
               AND apr.page_id = t_page_id
               AND apr.base_report_id = t_base_report_id
               AND apr.session_id = t_session_id;
            --
            SELECT col.report_label
                  ,col.column_alias
                  ,NULL BULK COLLECT
              INTO t_cols
              FROM apex_application_page_ir_col col
                  ,(SELECT substr(cols
                                 ,instr(sep || cols
                                       ,sep
                                       ,1
                                       ,LEVEL)
                                 ,instr(cols || sep
                                       ,sep
                                       ,1
                                       ,LEVEL) - instr(sep || cols
                                                      ,sep
                                                      ,1
                                                      ,LEVEL)) col
                          ,LEVEL so
                      FROM (SELECT t_apr.report_columns cols
                                  ,':' sep
                              FROM dual)
                    CONNECT BY LEVEL <=
                               length(cols) - NVL(length(REPLACE(cols
                                                                ,sep))
                                                 ,0) + 1) fil
             WHERE col.application_id = t_app_id
               AND col.page_id = t_page_id
               AND col.region_id = r_apr.region_id
               AND col.column_alias = fil.col
             ORDER BY fil.so;
            --
            FOR r_con IN (SELECT con.condition_sql
                                ,con.condition_type
                                ,con.condition_expr_type
                                ,con.condition_expression
                                ,con.condition_expression2
                            FROM apex_application_page_ir_cond con
                           WHERE con.application_id = t_app_id
                             AND con.page_id = t_page_id
                             AND con.report_id = t_apr.report_id
                             AND con.condition_enabled = 'Yes'
                             AND con.condition_type IN
                                 ('Filter'
                                 ,'Search'))
            LOOP
              IF (r_con.condition_type = 'Filter' AND
                 r_con.condition_sql IS NOT NULL)
              THEN
                IF r_con.condition_expr_type = 'ROW'
                THEN
                  t_where := t_where || chr(13) || 'and (' ||
                             r_con.condition_sql || ')';
                ELSE
                  IF instr(r_con.condition_sql
                          ,'#APXWS_EXPR#') > 0
                  THEN
                    r_con.condition_sql := REPLACE(r_con.condition_sql
                                                  ,'#APXWS_EXPR#'
                                                  ,'''' ||
                                                   REPLACE(r_con.condition_expression
                                                          ,''''
                                                          ,'''''') || '''');
                  END IF;
                  IF instr(r_con.condition_sql
                          ,'#APXWS_EXPR2#') > 0
                  THEN
                    r_con.condition_sql := REPLACE(r_con.condition_sql
                                                  ,'#APXWS_EXPR2#'
                                                  ,'''' ||
                                                   REPLACE(r_con.condition_expression2
                                                          ,''''
                                                          ,'''''') || '''');
                  END IF;
                  --
                  t_start   := 1;
                  t_val_ind := 1;
                  WHILE instr(r_con.condition_sql
                             ,'#APXWS_EXPR_VAL' || t_val_ind || '#') > 0
                  LOOP
                    t_end := instr(r_con.condition_expression || ','
                                  ,','
                                  ,t_start);
                    t_val := TRIM(substr(r_con.condition_expression
                                        ,t_start
                                        ,t_end - t_start));
                    IF t_val IS NULL
                    THEN
                      r_con.condition_sql := REPLACE(r_con.condition_sql
                                                    ,', #APXWS_EXPR_VAL' ||
                                                     t_val_ind || '#');
                      r_con.condition_sql := REPLACE(r_con.condition_sql
                                                    ,'#APXWS_EXPR_VAL' ||
                                                     t_val_ind || '# ,');
                    ELSE
                      r_con.condition_sql := REPLACE(r_con.condition_sql
                                                    ,'#APXWS_EXPR_VAL' ||
                                                     t_val_ind || '#'
                                                    ,'''' ||
                                                     REPLACE(t_val
                                                            ,''''
                                                            ,'''''') || '''');
                    END IF;
                    t_val_ind := t_val_ind + 1;
                    t_start   := t_end + 1;
                  END LOOP;
                  t_where := t_where || chr(13) || 'and ' ||
                             r_con.condition_sql;
                END IF;
              END IF;
              --
              IF r_con.condition_type = 'Search'
              THEN
                FOR r_sea IN (SELECT col.column_alias
                                FROM apex_application_page_ir_col col
                                    ,(SELECT substr(cols
                                                   ,instr(sep || cols
                                                         ,sep
                                                         ,1
                                                         ,LEVEL)
                                                   ,instr(cols || sep
                                                         ,sep
                                                         ,1
                                                         ,LEVEL) -
                                                    instr(sep || cols
                                                         ,sep
                                                         ,1
                                                         ,LEVEL)) col
                                            ,LEVEL so
                                        FROM (SELECT t_apr.report_columns cols
                                                    ,':' sep
                                                FROM dual)
                                      CONNECT BY LEVEL <= length(cols) -
                                                 NVL(length(REPLACE(cols
                                                                            ,sep))
                                                             ,0) + 1) fil
                               WHERE col.application_id = t_app_id
                                 AND col.page_id = t_page_id
                                 AND col.region_id = r_apr.region_id
                                 AND col.column_alias = fil.col
                                 AND col.allow_filtering = 'Yes'
                                 AND col.column_type IN
                                     ('STRING'
                                     ,'NUMBER'
                                     ,'CLOB'))
                LOOP
                  t_search := t_search || 'or instr( upper( "' ||
                              r_sea.column_alias || '" ), upper( ''' ||
                              REPLACE(r_con.condition_expression
                                     ,''''
                                     ,'''''') || ''' ) ) > 0 ';
                END LOOP;
                IF t_search IS NOT NULL
                THEN
                  t_where := t_where || chr(13) || 'and ( ' ||
                             ltrim(t_search
                                  ,'or') || ' )';
                END IF;
              END IF;
            END LOOP;
            --
            IF t_where IS NOT NULL
            THEN
              r_apr.region_source := 'select * from ( ' ||
                                     r_apr.region_source || ' ) where ' ||
                                     ltrim(t_where
                                          ,chr(13) || 'and');
            END IF;
            IF t_apr.sort_column_1 IS NOT NULL
            THEN
              r_apr.region_source := r_apr.region_source || ' order by "' ||
                                     t_apr.sort_column_1 || '" ' ||
                                     t_apr.sort_direction_1;
              IF t_apr.sort_column_2 IS NOT NULL
              THEN
                r_apr.region_source := r_apr.region_source || ', "' ||
                                       t_apr.sort_column_2 || '" ' ||
                                       t_apr.sort_direction_2;
              END IF;
              IF t_apr.sort_column_3 IS NOT NULL
              THEN
                r_apr.region_source := r_apr.region_source || ', "' ||
                                       t_apr.sort_column_3 || '" ' ||
                                       t_apr.sort_direction_3;
              END IF;
              IF t_apr.sort_column_4 IS NOT NULL
              THEN
                r_apr.region_source := r_apr.region_source || ', "' ||
                                       t_apr.sort_column_4 || '" ' ||
                                       t_apr.sort_direction_4;
              END IF;
              IF t_apr.sort_column_5 IS NOT NULL
              THEN
                r_apr.region_source := r_apr.region_source || ', "' ||
                                       t_apr.sort_column_5 || '" ' ||
                                       t_apr.sort_direction_5;
              END IF;
              IF t_apr.sort_column_6 IS NOT NULL
              THEN
                r_apr.region_source := r_apr.region_source || ', "' ||
                                       t_apr.sort_column_6 || '" ' ||
                                       t_apr.sort_direction_6;
              END IF;
            END IF;
          END;
        ELSE
          SELECT col.report_label
                ,col.column_alias
                ,NULL BULK COLLECT
            INTO t_cols
            FROM apex_application_page_ir_col col
           WHERE col.region_id = r_apr.region_id
           ORDER BY col.display_order;
        END IF;
      ELSE
        DECLARE
          t_use_filters VARCHAR2(1);
          t_pref        VARCHAR2(32767);
          t_order       VARCHAR2(32767);
        BEGIN
          IF p_use_filters
          THEN
            t_use_filters := 'Y';
            t_pref        := apex_util.get_preference('FSP' || t_app_id || '_P' ||
                                                      t_page_id || '_R' ||
                                                      r_apr.region_id ||
                                                      '_SORT'
                                                     ,v('APP_USER'));
            IF instr(t_pref
                    ,'fsp_sort_') = 1
            THEN
              t_order := substr(t_pref
                               ,10);
              IF instr(t_order
                      ,'desc') > 0
              THEN
                t_order := substr(t_order
                                 ,1
                                 ,instr(t_order
                                       ,'_') - 1);
                t_order := t_order || ' asc'; -- bug in APEX????
              ELSE
                t_order := t_order || ' desc';
              END IF;
              r_apr.region_source := 'select * from ( ' ||
                                     r_apr.region_source || ' ) order by ' ||
                                     REPLACE(substr(t_pref
                                                   ,10)
                                            ,'_'
                                            ,' ');
            END IF;
            IF t_order IS NOT NULL
            THEN
              r_apr.region_source := 'select * from ( ' ||
                                     r_apr.region_source || ' ) order by ' ||
                                     t_order;
            END IF;
          ELSE
            t_use_filters := 'N';
          END IF;
          apex_debug_message.log_message(r_apr.region_source
                                        ,p_level => 1);
          --
          SELECT col.heading
                ,col.column_alias
                ,NULL BULK COLLECT
            INTO t_cols
            FROM apex_application_page_rpt_cols col
           WHERE col.region_id = r_apr.region_id
             AND (NVL(col.include_in_export
                     ,'Yes') = 'Yes' OR t_use_filters = 'N')
           ORDER BY col.display_sequence;
        END;
      END IF;
      --
      IF p_filename = 'PAGE_NAME'
      THEN
        l_filename := r_apr.page_name;
      ELSIF p_filename = 'REGION_NAME'
      THEN
        l_filename := r_apr.region_name;
      END IF;

      --
      -- New sheet
      --
      IF p_sheetname
      THEN
        l_sheet_name := dbms_xmlgen.convert(translate(r_apr.region_name
                                                     ,'a/\[]*:?'
                                                     ,'a'));
      ELSE
        l_sheet_name := 'Sheet' || l_sheet_id;
      END IF;
      --
      as_xlsx.new_sheet(l_sheet_name);
      --
      -- get the sorting
      -- FSP<app>P<page>R<region?>SORT
      l_order_by := htmldb_util.get_preference('FSP' || t_app_id || '_P' ||
                                               t_page_id || '_R' ||
                                               r_apr.region_id || '_SORT'
                                              ,v('APP_USER'));
      l_order_by := REPLACE(l_order_by
                           ,'fsp_sort_'
                           ,'');
      l_order_by := REPLACE(l_order_by
                           ,'_'
                           ,' ');
      IF length(l_order_by) > 0
      THEN
        IF instr(l_order_by
                ,'desc') > 0
        THEN
          l_order_by := REPLACE(l_order_by
                               ,'desc'
                               ,'asc');
        ELSE
          l_order_by := l_order_by || ' desc';
        END IF;
        l_order_by := ' order by ' || l_order_by;
      END IF;

      --
      BEGIN
        t_val := apex_plugin_util.get_data2(r_apr.region_source || l_order_by
                                           ,1
                                           ,c_max_columns
                                           ,p_component_name => p_process.name);
      EXCEPTION
        WHEN OTHERS THEN
          apex_debug_message.log_message(r_apr.region_source
                                        ,p_level => 1);
          RAISE;
      END;
      --

      --
      FOR c IN 1 .. t_cols.count()
      LOOP
        IF r_apr.report_column_source_type = 'GENERIC_REPORT_COLUMNS'
        THEN
          t_cols(c).ind := to_number(substr(t_cols(c).alias
                                           ,4));
        ELSE
          FOR cc IN 1 .. t_val.count()
          LOOP
            IF t_val(cc).name = t_cols(c).alias
            THEN
              t_cols(c).ind := cc;
              EXIT;
            END IF;
          END LOOP;
        END IF;
      END LOOP;
      --
      -- HEADER

      IF p_headers
      THEN
        l_column := 0;
        FOR c IN 1 .. t_cols.count()
        LOOP
          -- if column number is NOT on the "exclude column list"
          IF INSTR(',' || p_exclude_column_list || ','
                  ,',' || c || ',') = 0
          THEN
            l_column := l_column + 1;
            IF t_val.exists(t_cols(c).ind)
            THEN
              IF t_val(t_cols(c).ind)
               .data_type IN
                  (apex_plugin_util.c_data_type_varchar2
                  ,apex_plugin_util.c_data_type_number
                  ,apex_plugin_util.c_data_type_date
                  ,apex_plugin_util.c_data_type_clob
                  ,apex_plugin_util.c_data_type_timestamp
                  ,apex_plugin_util.c_data_type_timestamp_tz
                  ,apex_plugin_util.c_data_type_timestamp_ltz)
              THEN
                as_xlsx.cell(l_column
                            ,1
                            ,t_cols     (c).label
                            ,p_alignment => as_xlsx.get_alignment(p_vertical   => 'bottom'
                                                                 ,p_horizontal => 'general')
                            ,p_fillId    => as_xlsx.get_fill(p_patternType => 'solid'
                                                            ,p_fgRGB       => 'DAEEF3')
                            ,p_fontId    => as_xlsx.get_font(p_name     => 'PT Sans'
                                                            ,p_fontsize => 10
                                                            ,p_bold     => TRUE)
                            ,p_borderId  => as_xlsx.get_border(p_top    => 'thin'
                                                              ,p_bottom => 'thin'
                                                              ,p_left   => 'thin'
                                                              ,p_right  => 'thin'));

                as_xlsx.set_column_width(p_col   => l_column
                                        ,p_width => 24);

              END IF;
            END IF;
          END IF;
        END LOOP;
      END IF;

      --
      -- DATA
      --
      FOR r IN 1 .. t_val(1).value_list.count()
      LOOP
        IF p_headers
        THEN
          l_row := r + 1;
        ELSE
          l_row := r;
        END IF;
        l_column := 0;
        FOR c IN 1 .. t_cols.count()
        LOOP
          -- if column number is NOT on the "exclude column list"
          IF INSTR(',' || p_exclude_column_list || ','
                  ,',' || c || ',') = 0
          THEN
            l_column := l_column + 1;
            IF t_val.exists(t_cols(c).ind)
            THEN
              CASE t_val(t_cols(c).ind).data_type
                WHEN apex_plugin_util.c_data_type_varchar2 THEN
                  as_xlsx.cell(l_column
                              ,l_row
                              ,t_val(t_cols(c).ind).value_list(r)
                               .varchar2_value);
                WHEN apex_plugin_util.c_data_type_clob THEN
                  as_xlsx.cell(l_column
                              ,l_row
                              ,substr(t_val(t_cols(c).ind).value_list(r)
                                      .clob_value
                                     ,1
                                     ,32767));
                WHEN apex_plugin_util.c_data_type_number THEN
                  as_xlsx.cell(l_column
                              ,l_row
                              ,t_val(t_cols(c).ind).value_list(r)
                               .number_value);
                WHEN apex_plugin_util.c_data_type_date THEN
                  as_xlsx.cell(l_column
                              ,l_row
                              ,to_char(t_val(t_cols(c).ind).value_list(r)
                                       .date_value
                                      ,'yyyy-mm-dd"T"hh24:mi:ss."000"'));
                WHEN apex_plugin_util.c_data_type_timestamp THEN
                  as_xlsx.cell(l_column
                              ,l_row
                              ,to_char(t_val(t_cols(c).ind).value_list(r)
                                       .timestamp_value
                                      ,'yyyy-mm-dd"T"hh24:mi:ss.ff3'));
                WHEN apex_plugin_util.c_data_type_timestamp_tz THEN
                  as_xlsx.cell(l_column
                              ,l_row
                              ,to_char(t_val(t_cols(c).ind).value_list(r)
                                       .timestamp_tz_value
                                      ,'yyyy-mm-dd"T"hh24:mi:ss.ff3'));
                WHEN apex_plugin_util.c_data_type_timestamp_ltz THEN
                  -- do something with SESSIONTIMEZONE ?????
                  as_xlsx.cell(l_column
                              ,l_row
                              ,to_char(t_val(t_cols(c).ind).value_list(r)
                                       .timestamp_ltz_value
                                      ,'yyyy-mm-dd"T"hh24:mi:ss.ff3'));
                ELSE
                  NULL;
              END CASE;
            END IF;
          END IF;
        END LOOP;
      END LOOP;
      --
      l_sheet_id := l_sheet_id + 1;
    END LOOP;
    --
    IF l_sheet_id = 1
    THEN
      as_xlsx.cell(1
                  ,1
                  ,'No Report Region found!');
    END IF;
    --
    l_excel := as_xlsx.finish(p_landscape => FALSE);
    --
    IF p_filename = 'USER'
    THEN
      l_filename := p_user_filename;
    END IF;
    IF substr(l_filename
             ,-5) = '.xlsx'
    THEN
      l_filename := substr(l_filename
                          ,1
                          ,length(l_filename) - 5);
    END IF;
    l_filename := NVL(l_filename
                     ,'apex');
    IF p_append_sysdate != 'NO'
    THEN
      l_filename := l_filename || '_' ||
                    to_char(SYSDATE
                           ,p_append_sysdate);
    END IF;
    l_filename := l_filename || '.xlsx';
    htp.flush;
    owa_util.mime_header(wwv_flow_utilities.get_excel_mime_type
                        ,FALSE);
    htp.print('Content-Length: ' || dbms_lob.getlength(l_excel));
    htp.print('Content-disposition: attachment; filename=' || l_filename || ';');
    htp.print('Content-Description: Generated by as_xslx');
    owa_util.http_header_close;
    wpg_docload.download_file(l_excel);
    apex_debug_message.log_message('done');
    RETURN NULL;
  EXCEPTION
    WHEN OTHERS THEN
      apex_debug_message.log_message(dbms_utility.format_error_stack
                                    ,p_level => 1);
      apex_debug_message.log_message(dbms_utility.format_error_backtrace
                                    ,p_level => 1);
      DECLARE
        t_rv apex_plugin.t_process_exec_result;
      BEGIN
        t_rv.success_message := 'Oops, something went wrong' || '<br/>' ||
                                dbms_utility.format_error_stack ||
                                '<br/><br/>';
        RETURN t_rv;
      END;
  END;

END;
 
