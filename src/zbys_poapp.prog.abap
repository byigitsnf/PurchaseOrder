*&---------------------------------------------------------------------*
*& Report ZBYS_POAPP
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zbys_poapp.

TABLES: ekko, " PO Başlığı
        ekpo, " PO Kalemleri
        lfa1, " Tedarikçi Genel Verileri
        adr6. " SMTP Adresleri

* Dahili Tablolar ve Çalışma Alanları Tanımlamaları

DATA: gt_ekko         TYPE STANDARD TABLE OF ekko,   " PO Başlıkları için dahili tablo
      gs_ekko         TYPE ekko,                     " PO Başlığı için çalışma alanı
      gt_ekpo         TYPE STANDARD TABLE OF ekpo,   " PO Kalemleri için dahili tablo
      gs_lfa1         TYPE lfa1,                     " Tedarikçi için çalışma alanı
      gs_adr6         TYPE adr6,                     " E-posta adresi için çalışma alanı
      gv_vendor_email TYPE adr6-smtp_addr.           " Tedarikçi E-posta adresini tutacak değişken

* ALV ve XLS Ek için birleşik veri yapısı ve tablosu
* PO başlık ve kalemlerinden ilgili alanları içerecek

TYPES: BEGIN OF ty_po_data,
         ebeln      TYPE ekko-ebeln, " Sipariş Numarası
         aedat      TYPE ekko-aedat, " Belge Tarihi
         lifnr      TYPE ekko-lifnr, " Tedarikçi Numarası
         name1      TYPE lfa1-name1, " Tedarikçi Adı
         ebelp      TYPE ekpo-ebelp, " Kalem Numarası
         matnr      TYPE ekpo-matnr, " Malzeme Numarası
         menge      TYPE ekpo-menge, " Miktar
         meins      TYPE ekpo-meins, " Birim
         netpr      TYPE ekpo-netpr, " Net Fiyat
         netwr_item TYPE ekpo-netwr, " Kalem Net Değeri
         waers      TYPE ekko-waers, " Para Birimi
       END OF ty_po_data.

DATA: gt_po_data TYPE STANDARD TABLE OF ty_po_data, " ALV ve XLS için ana veri tablosu
      gs_po_data TYPE ty_po_data.                   " Çalışma alanı

* ALV için gerekli değişkenler

DATA: gt_fieldcatalog TYPE slis_t_fieldcat_alv, " Field catalog tablosu
      gs_fieldcatalog TYPE slis_fieldcat_alv.   " Field catalog çalışma alanı

* Mail gönderme için gerekli değişkenler (CL_BCS)

DATA: go_bcs       TYPE REF TO cl_bcs,
      go_doc_bcs   TYPE REF TO cl_document_bcs,   " E-posta doküman objesi
      go_recipient TYPE REF TO if_recipient_bcs,
      gt_soli      TYPE TABLE OF soli,            " E-posta gövdesi için (SOLI formatı)
      gv_content   TYPE string,                   " E-posta gövdesi için (STRING formatı - Metin)
      gv_status    TYPE bcs_rqst.                 " Durum değişkeni

* XLS Ek için gerekli değişkenler (Sizin çalışan yapı)

DATA: gv_attachment_size TYPE sood-objlen,       " Ek boyutu
      gt_att_content_hex TYPE solix_tab,         " Ek içeriği (binary - SOLIX_TAB)
      gv_att_content     TYPE string,            " Ek içeriği için string (TSV formatında)
      gv_att_line        TYPE string.            " Ek dosyasının tek bir satırı için (TSV formatında)

SELECTION-SCREEN BEGIN OF BLOCK selection_block WITH FRAME TITLE TEXT-001. " Seçim ekranı bloğu

  PARAMETERS: s_ebeln TYPE ekko-ebeln OBLIGATORY. " PO Numarası, zorunlu alan
  PARAMETERS: p_mail AS CHECKBOX DEFAULT 'X'.     " Mail Gönder checkbox'ı, varsayılan olarak seçili

  PARAMETERS: p_email TYPE adr6-smtp_addr. " Test Mail Adresi

SELECTION-SCREEN END OF BLOCK selection_block.

* Metin Sembolleri
* TEXT-001: 'Sipariş E-Posta Gönderimi'
* TEXT-002: 'PO Numarası & Bulunamadı.'
* TEXT-003: 'PO & için Tedarikçi Bulunamadı.'
* TEXT-004: 'Tedarikçi & için E-posta Adresi Bulunamadı.'
* TEXT-005: 'Mail gönderme seçeneği işaretli değil. İşlem yapılmadı.'
* TEXT-006: 'Sipariş & için E-posta başarıyla gönderildi.'
* TEXT-007: 'Sipariş & için E-posta gönderilirken bir hata oluştu.'
* TEXT-012: 'E-posta Dokümanı Oluşturulurken Hata (SY-SUBRC &).'
* TEXT-013: 'Alıcı Adresi Oluşturulurken Hata (SY-SUBRC &).'
* TEXT-014: 'Send Request Oluşturulurken Hata (SY-SUBRC &).'
* TEXT-015: 'Doküman Send Request\'e Eklenirken Hata (SY-SUBRC &).'
* TEXT-016: 'Alıcı Send Request\'e Eklenirken Hata (SY-SUBRC &).'
* TEXT-017: 'Durum Ayarlanırken Hata (SY-SUBRC &).'
* TEXT-018: 'E-posta Gönderme Başlatılırken Hata.'
* TEXT-024: 'XLS Ek İçeriği Dönüşümünde Hata (SY-SUBRC &).'
* TEXT-025: 'Mail Gönderilecek Test Adresi'

START-OF-SELECTION.

  " 1. EKKO tablosundan PO Başlık verileri
  SELECT SINGLE *
    INTO gs_ekko
    FROM ekko
    WHERE ebeln = s_ebeln.

  IF sy-subrc <> 0.
    MESSAGE s002(00) WITH s_ebeln DISPLAY LIKE 'E'.
    EXIT.
  ENDIF.

  " 2. EKPO tablosundan PO Kalem verileri
  SELECT *
    INTO TABLE gt_ekpo
    FROM ekpo
    WHERE ebeln = s_ebeln.

  " 3. LFA1 tablosından Tedarikçi verileri
  SELECT SINGLE *
    INTO gs_lfa1
    FROM lfa1
    WHERE lifnr = gs_ekko-lifnr.

  IF sy-subrc <> 0.
    MESSAGE s003(00) WITH s_ebeln DISPLAY LIKE 'E'.
    EXIT.
  ENDIF.

  " 4. ADR6 tablosundan Tedarikçinin E-posta Adresini çek (Bilgi amaçlı, kullanılmıyor)
  SELECT SINGLE *
    INTO gs_adr6
    FROM adr6
    WHERE addrnumber = gs_lfa1-adrnr.
  " AND   smtp_valid = 'X'.

  gv_vendor_email = gs_adr6-smtp_addr.

  " ALV Veri Tablosunu Doldurma (EKKO ve EKPO verileri birleştirilir)

  CLEAR gt_po_data.
  IF gt_ekpo IS NOT INITIAL.
    LOOP AT gt_ekpo ASSIGNING FIELD-SYMBOL(<fs_ekpo>).
      CLEAR gs_po_data.

      " Başlık verileri
      gs_po_data-ebeln = gs_ekko-ebeln.
      gs_po_data-aedat = gs_ekko-aedat.
      gs_po_data-lifnr = gs_ekko-lifnr.
      gs_po_data-name1 = gs_lfa1-name1.
      gs_po_data-waers = gs_ekko-waers.

      " Kalem verileri
      gs_po_data-ebelp = <fs_ekpo>-ebelp.
      gs_po_data-matnr = <fs_ekpo>-matnr.
      gs_po_data-menge = <fs_ekpo>-menge.
      gs_po_data-meins = <fs_ekpo>-meins.
      gs_po_data-netpr = <fs_ekpo>-netpr.
      gs_po_data-netwr_item = <fs_ekpo>-netwr.

      APPEND gs_po_data TO gt_po_data.

    ENDLOOP.
  ENDIF.

  " ALV Ekranını Gösterme (Veriyi görmek için)

  IF gt_po_data IS NOT INITIAL.

    PERFORM build_fieldcatalog USING gt_fieldcatalog.

    CALL FUNCTION 'REUSE_ALV_GRID_DISPLAY'
      EXPORTING
        i_callback_program = sy-repid
        it_fieldcat        = gt_fieldcatalog
        i_default          = 'X'
        i_save             = 'X'
      TABLES
        t_outtab           = gt_po_data
      EXCEPTIONS
        program_error      = 1
        joining_disabled   = 2
        ---------          = 3.
    IF sy-subrc <> 0.
      MESSAGE i000(00) WITH 'ALV Gösterilirken Hata Oluştu.'.
    ENDIF.
  ELSE.
    MESSAGE i000(00) WITH 'PO Kalem Verisi Bulunamadı.'.
  ENDIF.

  " Mail Gönderme Kontrolü

  IF p_mail IS NOT INITIAL. " p_mail checkbox'ı işaretliyse devam et

    " Test Mail Adresi Kontrolü

    IF p_email IS INITIAL.
      MESSAGE e000(00) WITH 'Test Mail Adresi Boş Bırakılamaz.'.
      EXIT.
    ENDIF.

    " E-posta İçeriğini Oluşturma (Düz Metin Formatında - CONCATENATE Kullanarak)

    CLEAR gv_content.
    CONCATENATE 'Sayın Yetkili,' cl_abap_char_utilities=>cr_lf INTO gv_content.
    CONCATENATE gv_content 'Aşağıdaki Siparişinizin Bilgileri Ekte Yer Almaktadır:' cl_abap_char_utilities=>cr_lf INTO gv_content.
    CONCATENATE gv_content cl_abap_char_utilities=>cr_lf INTO gv_content. " Boş satır

    CONCATENATE gv_content 'Siparis Numarasi: ' gs_ekko-ebeln cl_abap_char_utilities=>cr_lf INTO gv_content.

    DATA: lv_aedat_char TYPE char10.
    WRITE gs_ekko-aedat TO lv_aedat_char.

    CONCATENATE gv_content 'Belge Tarihi: ' lv_aedat_char cl_abap_char_utilities=>cr_lf INTO gv_content.

    CONCATENATE gv_content 'Tedarikci: ' gs_lfa1-name1 cl_abap_char_utilities=>cr_lf INTO gv_content.

    IF gt_po_data IS NOT INITIAL.
      CONCATENATE gv_content 'PO Kalem Bilgileri ekteki XLS dosyasinda detaylandirilmistir.' cl_abap_char_utilities=>cr_lf INTO gv_content.
    ENDIF.

    CONCATENATE gv_content cl_abap_char_utilities=>cr_lf INTO gv_content. " Boş satır
    CONCATENATE gv_content 'Bilgilerinize sunulur.' cl_abap_char_utilities=>cr_lf INTO gv_content.
    CONCATENATE gv_content 'Saygılarımızla,' cl_abap_char_utilities=>cr_lf INTO gv_content.
    CONCATENATE gv_content 'Satin Alma Departmani' cl_abap_char_utilities=>cr_lf INTO gv_content.

    " Metin içeriği SOLI tablosuna dönüştür

    gt_soli = cl_document_bcs=>string_to_soli( gv_content ).

    " XLS Ek İçeriğini Oluşturma (ALV Verisi - TSV Formatında)

    CLEAR gv_att_content.
    CLEAR gt_att_content_hex.
    CLEAR gv_attachment_size.

    " Başlık satırını oluşturma

    CONCATENATE 'Siparis No'
                'Belge Tarihi'
                'Tedarikci No'
                'Tedarikci Adi'
                'Kalem No'
                'Malzeme No'
                'Miktar'
                'Birim'
                'Net Fiyat'
                'Kalem Net Degeri'
                'Para Birimi'
                INTO gv_att_content
                SEPARATED BY cl_abap_char_utilities=>horizontal_tab.

    IF gt_po_data IS NOT INITIAL.

      " İlk başlık satırından sonra newline karakteri ekle

      CONCATENATE gv_att_content cl_abap_char_utilities=>newline INTO gv_att_content.

      " ALV veri tablosundaki verileri ekle

      LOOP AT gt_po_data ASSIGNING FIELD-SYMBOL(<fs_po_data>).
        DATA: lv_ebelp_char TYPE char6.
        lv_ebelp_char = <fs_po_data>-ebelp.
        CONDENSE lv_ebelp_char NO-GAPS.

        DATA: lv_menge_char TYPE char17.
        WRITE <fs_po_data>-menge TO lv_menge_char.
        CONDENSE lv_menge_char NO-GAPS.

        DATA: lv_meins_char TYPE char3.
        WRITE <fs_po_data>-meins TO lv_meins_char.

        DATA: lv_netpr_char TYPE char17.
        WRITE <fs_po_data>-netpr TO lv_netpr_char CURRENCY <fs_po_data>-waers.
        CONDENSE lv_netpr_char NO-GAPS.

        DATA: lv_netwr_item_char TYPE char17.
        WRITE <fs_po_data>-netwr_item TO lv_netwr_item_char CURRENCY <fs_po_data>-waers.
        CONDENSE lv_netwr_item_char NO-GAPS.


        " Veri satırını oluştur

        CONCATENATE <fs_po_data>-ebeln
                    <fs_po_data>-aedat
                    <fs_po_data>-lifnr
                    <fs_po_data>-name1
                    lv_ebelp_char
                    <fs_po_data>-matnr
                    lv_menge_char
                    lv_meins_char
                    lv_netpr_char
                    lv_netwr_item_char
                    <fs_po_data>-waers
               INTO gv_att_line
               SEPARATED BY cl_abap_char_utilities=>horizontal_tab.

        " Ana ek içeriği stringine satırı ekle

        CONCATENATE gv_att_content cl_abap_char_utilities=>newline gv_att_line INTO gv_att_content.

      ENDLOOP.
    ENDIF.

    " XLS içeriği stringini binary (SOLIX_TAB) formatına dönüştür

    CLEAR sy-subrc.
    CALL METHOD cl_bcs_convert=>string_to_solix
      EXPORTING
        iv_string   = gv_att_content
        iv_codepage = '4103' " UTF-8
        iv_add_bom  = 'X'
      IMPORTING
        et_solix    = gt_att_content_hex
        ev_size     = gv_attachment_size
      EXCEPTIONS
        OTHERS      = 1.

    IF sy-subrc <> 0.
      MESSAGE e024(00) WITH sy-subrc.
      " Dönüşüm hatası varsa eki gönderme
      CLEAR gt_att_content_hex.
      CLEAR gv_attachment_size.
    ENDIF.

    " E-postayı Gönderme (CL_BCS Kullanarak - Düz Metin Gövde ve XLS Ek)

    CLEAR sy-subrc.

    " E-posta Doküman objesini oluştur (Düz Metin Gövde ve Konu)

    CALL METHOD cl_document_bcs=>create_document
      EXPORTING
        i_type    = 'RAW' " düz metin için kullanılır
        i_text    = gt_soli          " SOLI formatındaki gövde içeriği
        i_subject = 'Siparis Bilgileri: ' && gs_ekko-ebeln " E-posta Konusu
      RECEIVING
        result    = go_doc_bcs
      EXCEPTIONS
        OTHERS    = 1.

    IF sy-subrc <> 0.
      MESSAGE e012(00) WITH sy-subrc.
      EXIT.
    ENDIF.

    " XLS Ekini Ekle (Ek içeriği başarıyla oluşturulduysa)

    IF gt_att_content_hex IS NOT INITIAL.

      " Eki Doküman objesine ekle

      CLEAR sy-subrc.

      CALL METHOD go_doc_bcs->add_attachment
        EXPORTING
          i_attachment_type    = 'xls'
          i_attachment_subject = 'PO_Kalemleri_' && gs_ekko-ebeln && '.xls'
          i_attachment_size    = gv_attachment_size
          i_att_content_hex    = gt_att_content_hex
        EXCEPTIONS
          OTHERS               = 1.

      IF sy-subrc <> 0.
        MESSAGE e000(00) WITH 'Eke XLS Eklenirken Hata (SY-SUBRC &)'.
      ENDIF.

    ENDIF.


    " Alıcı objesini oluştur (SMTP Adres)

    CLEAR sy-subrc.
    CALL METHOD cl_cam_address_bcs=>create_internet_address
      EXPORTING
        i_address_string = p_email " Test Mail Adresi parametresi kullanılıyor
      RECEIVING
        result           = go_recipient
      EXCEPTIONS
        OTHERS           = 1.

    IF sy-subrc <> 0.
      MESSAGE e013(00) WITH sy-subrc.
      EXIT.
    ENDIF.

    " Send Request objesini oluştur

    CLEAR sy-subrc.
    CALL METHOD cl_bcs=>create_persistent
      RECEIVING
        result = go_bcs
      EXCEPTIONS
        OTHERS = 1.

    IF sy-subrc <> 0.
      MESSAGE e014(00) WITH sy-subrc.
      EXIT.
    ENDIF.

    " Dokümanı Send Request'e ekle

    CLEAR sy-subrc.
    CALL METHOD go_bcs->set_document
      EXPORTING
        i_document = go_doc_bcs
      EXCEPTIONS
        OTHERS     = 1.

    IF sy-subrc <> 0.
      MESSAGE e015(00) WITH sy-subrc.
      EXIT.
    ENDIF.

    " Alıcıyı Send Request'e ekle

    CLEAR sy-subrc.
    CALL METHOD go_bcs->add_recipient
      EXPORTING
        i_recipient = go_recipient
      EXCEPTIONS
        OTHERS      = 1.

    IF sy-subrc <> 0.
      MESSAGE e016(00) WITH sy-subrc.
      EXIT.
    ENDIF.

    " Durum özniteliklerini ayarlama

    gv_status = 'N'.
    CLEAR sy-subrc.
    CALL METHOD go_bcs->set_status_attributes
      EXPORTING
        i_requested_status = gv_status
      EXCEPTIONS
        OTHERS             = 1.

    IF sy-subrc <> 0.
      MESSAGE e017(00) WITH sy-subrc.
    ENDIF.

    " E-postayı Gönder

    DATA(lv_send_ok) = abap_false.
    CLEAR sy-subrc.
    CALL METHOD go_bcs->send
      RECEIVING
        result = lv_send_ok
      EXCEPTIONS
        OTHERS = 1.

    IF sy-subrc <> 0.
      MESSAGE e018(00) WITH sy-subrc.
    ELSEIF lv_send_ok = abap_true.
      COMMIT WORK.
      MESSAGE s006(00) WITH s_ebeln DISPLAY LIKE 'S'.
    ELSE.
      MESSAGE s007(00) WITH s_ebeln DISPLAY LIKE 'E'.
    ENDIF.

  ELSE. " p_mail checkbox'ı işaretli değilse
    MESSAGE s005(00) DISPLAY LIKE 'I'.
  ENDIF. " IF p_mail IS NOT INITIAL.

END-OF-SELECTION.

*&---------------------------------------------------------------------*
*&      Form  BUILD_FIELDCATALOG
*&---------------------------------------------------------------------*

FORM build_fieldcatalog USING pt_fieldcatalog TYPE slis_t_fieldcat_alv.

  CLEAR pt_fieldcatalog.

  DEFINE add_field.
    CLEAR gs_fieldcatalog.
    gs_fieldcatalog-fieldname = &1.
    gs_fieldcatalog-tabname = 'GT_PO_DATA'.
    gs_fieldcatalog-col_pos = &2.
    gs_fieldcatalog-seltext_l = &3.
    gs_fieldcatalog-outputlen = &4.
    APPEND gs_fieldcatalog TO pt_fieldcatalog.
  END-OF-DEFINITION. " *** BURASI MAKRO TANIMINI KAPATIR ***

  " ALV tablosundaki alanlara göre Field Catalog oluşturma

  add_field 'EBELN' 1 'Sipariş Numarası' 10.
  add_field 'AEDAT' 2 'Belge Tarihi' 10.
  add_field 'LIFNR' 3 'Tedarikçi No' 10.
  add_field 'NAME1' 4 'Tedarikçi Adı' 30.
  add_field 'EBELP' 5 'Kalem No' 6.
  add_field 'MATNR' 6 'Malzeme No' 18.
  add_field 'MENGE' 7 'Miktar' 17.
  add_field 'MEINS' 8 'Birim' 3.
  add_field 'NETPR' 9 'Net Fiyat' 17.
  add_field 'NETWR_ITEM' 10 'Kalem Değeri' 17.
  add_field 'WAERS' 11 'Para Birimi' 5.

ENDFORM.
