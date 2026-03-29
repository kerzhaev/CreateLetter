# Template Placeholders

CreateLetter replaces the following placeholders in Word templates.

## Preferred Placeholders

- `RecipientName`
- `RecipientAddress`
- `OutgoingNumber`
- `OutgoingDate`
- `ExecutorName`
- `ExecutorPhone`
- `LetterText`
- `AttachmentsList`

## Legacy Russian Placeholders Also Supported

- `НаименованиеПолучателя`
- `АдресПолучателя`
- `НомерИсходящего`
- `ДатаИсходящего`
- `ИсполнительФИО`
- `ТелефонИсполнителя`
- `ТекстПисьма`
- `СписокПриложений`

## Notes

- Use plain text placeholders inside the `.docx` template body.
- The preferred set is the English placeholder set above because it is stable in source control.
- Existing Russian templates continue to work because the workbook now replaces both preferred and legacy placeholder names.
