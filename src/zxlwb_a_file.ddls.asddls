@EndUserText.label: 'XLWB: rendered file (action result)'
define root abstract entity ZXLWB_A_FILE
{
  @EndUserText.label: 'File name'
  FileName      : abap.char(128);
  @EndUserText.label: 'MIME type'
  MimeType      : abap.char(128);
  @EndUserText.label: 'Content (base64)'
  ContentBase64 : abap.string(0);
}
