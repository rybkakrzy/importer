# language: pl
Właściwość: Documents API — operacje na dokumentach
  Jako klient API
  Chcę zarządzać dokumentami przez REST API
  Aby integrować D2 Tools z innymi systemami

  # ────────────────────────────────────────────
  # Nowy dokument
  # ────────────────────────────────────────────

  @api @smoke
  Scenariusz: Pobranie nowego dokumentu
    Kiedy wysyłam GET na "/document/new"
    Wtedy status odpowiedzi to 200
    I odpowiedź zawiera pole "html"

  # ────────────────────────────────────────────
  # Szablony
  # ────────────────────────────────────────────

  @api @smoke
  Scenariusz: Pobranie listy szablonów
    Kiedy wysyłam GET na "/document/templates"
    Wtedy status odpowiedzi to 200
    I odpowiedź jest listą

  @api @regression
  Szablon scenariusza: Pobranie konkretnego szablonu
    Kiedy wysyłam GET na "/document/templates/<template_id>"
    Wtedy status odpowiedzi to 200
    I odpowiedź zawiera pole "html"

    Przykłady:
      | template_id |
      | blank       |
      | letter      |
      | report      |
      | cv          |

  # ────────────────────────────────────────────
  # Barcode API
  # ────────────────────────────────────────────

  @api @smoke
  Scenariusz: Pobranie obsługiwanych typów kodów
    Kiedy wysyłam GET na "/barcode/types"
    Wtedy status odpowiedzi to 200
    I odpowiedź jest listą

  @api @regression
  Szablon scenariusza: Generowanie kodu kreskowego przez API
    Kiedy wysyłam POST na "/barcode/generate" z danymi:
      | content      | barcodeType   | width | height |
      | <content>    | <barcodeType> | 300   | 300    |
    Wtedy status odpowiedzi to 200
    I odpowiedź zawiera pole "base64Image"

    Przykłady:
      | content             | barcodeType |
      | https://example.com | QRCode      |
      | ABC-12345           | Code128     |
      | 5901234123457       | EAN13       |

  # ────────────────────────────────────────────
  # Zapis dokumentu
  # ────────────────────────────────────────────

  @api @regression
  Scenariusz: Zapisanie dokumentu jako DOCX
    Kiedy wysyłam POST na "/document/save" z danymi:
      | html                        | originalFileName |
      | <p>Test dokumentu</p>       | test.docx        |
    Wtedy status odpowiedzi to 200
    I odpowiedź ma content-type "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
