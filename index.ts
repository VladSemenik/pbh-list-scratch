import puppeteer from "puppeteer";
import xlsx from "xlsx";

function isLink(s: string) {
  return s.includes("http://") || s.includes("https://");
}

function isMail(s: string) {
  return s.includes("mailto:");
}

(async () => {
  const browser = await puppeteer.launch({
    product: "chrome",
    headless: false,
  });
  const page = await browser.newPage();

  await page.goto(
    "https://www.gov.pl/web/poland-businessharbour-en/itspecialist"
  );

  const container = await page.$$("#main-content >>> details");

  let c = 0;
  const arr: {
    name?: string;
    link?: string;
    mail?: string;
  }[] = [];
  for await (const el of container) {
    ++c;
    const as = await el.$$("a");
    const summary = await el.$("summary");

    const line: {
      name?: string;
      link?: string;
      mail?: string;
    } = {};

    line.name =
      (await summary?.getProperty("innerText"))
        ?.toString()
        .replace("JSHandle:", "") ?? "";

    for await (const a of as) {
      const s = (await a.getProperty("href"))
        .toString()
        .replace("JSHandle:", "");

      if (isLink(s)) {
        line.link = line.link ? line.link + " " + s : s;
      } else if (isMail(s)) {
        line.mail = line.mail ? line.mail + " " + s : s;
      }
    }

    arr.push(line);
  }

  const workbook = xlsx.utils.book_new();
  let worksheet = xlsx.utils.aoa_to_sheet([]);
  xlsx.utils.book_append_sheet(workbook, worksheet);
  xlsx.utils.sheet_add_aoa(worksheet, [["Name", "Link", "Mail"]], {
    origin: "A1",
  });
  const formatedArr = arr.map((e) => [e.name, e.link, e.mail]);
  xlsx.utils.sheet_add_aoa(worksheet, formatedArr, { origin: "A2" });

  xlsx.writeFile(workbook, "pbh-list.xlsx");

  const workbook1 = xlsx.utils.book_new();
  let worksheet1 = xlsx.utils.aoa_to_sheet([]);
  xlsx.utils.book_append_sheet(workbook1, worksheet1);
  xlsx.utils.sheet_add_aoa(worksheet1, [["Name", "Link", "Mail"]], {
    origin: "A1",
  });
  const formatedArr1 = arr
    .filter((e) => e.link)
    .map((e) => [e.name, e.link, e.mail]);
  xlsx.utils.sheet_add_aoa(worksheet1, formatedArr1, { origin: "A2" });

  xlsx.writeFile(workbook1, "pbh-list-link-required.xlsx");

  await page.close();
  await browser.close();

  console.log("Done!");
})();
