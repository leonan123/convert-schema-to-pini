import * as fs from "fs"
import path from "path"
import { fileURLToPath } from "url"
import * as XLSX from "xlsx/xlsx.mjs"
XLSX.set_fs(fs)

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

const file = XLSX.readFile(path.join(__dirname, "/src/schemas/schema.xlsx"))

file.SheetNames.shift() // remove versao 2.00

const sheets = file.SheetNames

for (let schemaVersionIndex = 0; schemaVersionIndex < sheets.length; schemaVersionIndex++) {
  const data = XLSX.utils.sheet_to_json(file.Sheets[sheets[schemaVersionIndex]])

  const columns = [
    "CteValCodigo",
    "CteValTag",
    "CteValTagDescricao",
    "CteValRegEx",
    "CteValCodMsgErro",
    "CteValMsgErro",
    "CteValTagPai",
  ]

  let str = ""

  for (const row of data) {
    let columnsCount = 0
    for (const column of columns) {
      columnsCount++
      if (!row[column]) {
        row[column] = ""
      }
      
      str += `&${column} = ${typeof row[column] == "string" ? `'${row[column]}'` : row[column]}\n`

      if (columnsCount === columns.length) {
        str += "\n"
        columnsCount = 0
      }
    }
  }

  fs.writeFileSync(path.join(__dirname, `/src/schemas/converted/schema${sheets[schemaVersionIndex]}.txt`), str)
}
