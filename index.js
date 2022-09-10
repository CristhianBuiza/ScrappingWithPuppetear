const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");

const saveExcel = (data) => {
  const workbook = new ExcelJS.Workbook();
  const fileName = "Lista-Usuarios.xlsx";
  const sheet = workbook.addWorksheet("Resultados");
  const reColumns = [
    {
      header: "DNI",
      key: "dni",
    },
    {
      header: "Codigo_Servicio",
      key: "cod",
    },
  ];
  sheet.columns = reColumns;
  sheet.addRows(data);
  workbook.xlsx
    .writeFile(fileName)
    .then((e) => {
      console.log("Creado exitosamente");
    })
    .catch(() => {
      console.log("error");
    });
};

let data = [];
(async () => {
  const browser = await puppeteer.launch({ headless: false });
  const context = await browser.createIncognitoBrowserContext();
  const page = await context.newPage();
  loadUrl(
    page,
    "https://visorclientes.movistar.com.pe/login",
    context,
    browser
  );
})();
async function loadUrl(page, url, context, browser) {
  const userName = "";
  const password = "";
  const initialNumber = "";
  const finalNumber = "";
  await page.goto(url);
  await page.waitForSelector("#contenedor-botones", { visible: true });
  await page.click("#contenedor-botones > .tlf-bold-20");
  await page.waitForSelector("#signInName");
  await page.type("#signInName", userName);
  await page.type("#password", password);
  await page.waitForTimeout(1000);
  await page.click("#continue");
  // searchTerm
  try {
    for (let i = initialNumber; i <= finalNumber; i++) {
      await page.waitForSelector(".searchTerm", { visible: true });
      await page.type(".searchTerm", String(i));
      await page.click(".searchButton");
      await page.waitForTimeout(1500);
      // search terminated
      try {
        await page.waitForSelector("#close-modal-message", {
          visible: true,
          timeout: 3000,
        });
        await page.click("#close-modal-message");
        await page.waitForTimeout(5000);
      } catch (e) {}
      const isActived = await page.evaluate(() => {
        try {
          let namePlan = document.querySelector("#namePlan").innerHTML;
        let isActive = document.querySelector(
          ".sidebar__status-information > .color-visor-white"
        ).innerHTML;
        if (namePlan == "TV ESTANDAR DIGITAL" && isActive == "Activo") {
          return true;
        } else return false;
        } catch (error) {
          return false;
        }
      });

      if (isActived == true) {
        const nroServicio = await page.$eval(
          "div.sidebar__service--code-service > div:nth-child(2)",
          (el) => el.textContent
        );
        await page.click(".toolbar-item-icon");
        const DNI = await page.$eval(
          "div.client-space > div > p:nth-child(2)",
          (el) => el.textContent
        );
        data.push({
          dni: DNI,
          cod: nroServicio,
        });
        await page.click("[text='Cerrar atención']");
        await page.waitForTimeout(1000);
        await page.click("[label='Salir de la atención']");
        await page.waitForTimeout(1000);
      } else if (isActived == false) {
        await page.click("[text='Cerrar atención']");
        await page.waitForTimeout(1000);
        await page.click("[label='Salir de la atención']");
        //[text='Cerrar atención']
        await page.waitForTimeout(1000);
      }
    }
  } catch (error) {
    console.log(error);
  } finally {
    saveExcel(data);
    await context.close();
    await browser.close();
  }
}
