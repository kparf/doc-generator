import { PatchType, patchDocument, TextRun } from "docx";
import * as fs from "fs";
import { resolve, join } from "path";
import { program } from "commander";
import { parse } from "csv-parse";
import { Console } from "console";

const DEFAULT_OUTPUT_FOLDER = "./output";

program
  .name("Doc generator")
  .option("--kyes <string>", "path to csv file with kyes")
  .option("--out <string>", "path to csv file with kyes")
  .argument("<string>", "template file")
  .action(async (str, options) => {
    console.log("Start generation");

    console.log(str);
    console.log(options.kyes);

    const keysFilePath = resolve(options.kyes);
    const templateFilePath = resolve(str);
    const template = fs.readFileSync(templateFilePath);
    const keys = fs.readFileSync(keysFilePath, { encoding: "utf8" });

    const parser = fs.createReadStream(keysFilePath, { encoding : "utf8" }).pipe(
      parse({
        encoding: "utf8",
        delimiter: ",",
        columns: true,
        relax_quotes: true,
        escape: "\\",
        ltrim: true,
        rtrim: true,
      })
    );

    let index = 0;
    for await (const record of parser) {
      index++
      let patches: any = {}
      for (let prop in record) {
        patches[prop] = {
          type: PatchType.PARAGRAPH,
          children: [new TextRun(record[prop])],
        }
      }
      const resultDoc = await patchDocument(template, {
        patches
      });
  
      if (!fs.existsSync(resolve(DEFAULT_OUTPUT_FOLDER))) {
        fs.mkdirSync(resolve(DEFAULT_OUTPUT_FOLDER));
      }

      fs.writeFileSync(
        join(resolve(DEFAULT_OUTPUT_FOLDER), `doc-${index}.docx`),
        resultDoc
      );
    }
  });

program.parse();
