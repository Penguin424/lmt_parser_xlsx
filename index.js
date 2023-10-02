const Excel = require("exceljs");
const fs = require("fs");

String.prototype.replaceAll = function (search, replacement) {
  var target = this;
  return target.replace(new RegExp(search, "g"), replacement);
};

const main = async () => {
  const workbookF = new Excel.Workbook();

  const workbook = await workbookF.xlsx.readFile("./ON DOCK 2023.xlsx");
  const workSheet = workbook.worksheets[0];

  for (const image of workSheet.getImages()) {
    const sku = workSheet.getCell(`I${image.range.tl.nativeRow + 1}`);
    const description = workSheet.getCell(`I${image.range.tl.nativeRow + 1}`);
    const measurments = workSheet.getCell(`J${image.range.tl.nativeRow + 1}`);

    const img = workbook.model.media.find((m) => m.index === image.imageId);
    const measurmentsArray =
      measurments.value === null
        ? "0X0X0".split("X")
        : measurments.value
            .toString()
            .toUpperCase()
            .replaceAll(`"`, " ")
            .split("X");

    const l = parseFloat(measurmentsArray[2].trim());
    const w = parseFloat(measurmentsArray[1].trim());
    const h = parseFloat(measurmentsArray[0].trim());

    const skuparser =
      sku.value === null
        ? `NO SKU IN GOOGLE SHEET ${image.range.tl.nativeRow + 1}`
        : sku.value.toString().replaceAll("/", " ");

    // fs.writeFileSync(`./images/${skuparser}.${img.extension}`, img.buffer);
    console.log(measurmentsArray);
    const data = {
      name: skuparser,
      price: 0.1,
      sku: skuparser,
      description: description === null ? "N/A" : description,
      high: parseFloat(measurmentsArray[2].trim()),
      width: parseFloat(measurmentsArray[1].trim()),
      long: parseFloat(measurmentsArray[0].trim()),
      weight: parseFloat(),
      wholesale: 0.1,
      retail: 0.1,
      decorator: 0.1,
      dropship: 0.1,
      images: [159],
      supplier: 2,
      catalog: 2,
      categories: 1,
      isAvailable: true,
      inventory: 21,
      creatorat: 1,
      journey: 0,
      ordereded: 0,
      stock: 0,
      paragraphs: 0,
    };

    console.log(data);
  }
};

main();
