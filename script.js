const url = "https://restcountries.com/v3.1/all";
const axios = require("axios").default;
const xl = require("excel4node");

function getCountries() {
  axios
    .get(url)
    .then((res) => {
      const data = res.data;
      let infoCountries = [];

      let country = {
        name: "",
        capital: "",
        area: "",
        currencies: "",
      };

      data.forEach(function (value, key) {
        country.name = value.name.common;

        value.capital?.[0] !== undefined
          ? (country.capital = value.capital[0])
          : (country.capital = "-");

        value.area !== undefined
          ? (country.area = value.area)
          : (country.area = "-");

        country.currencies = Object.keys(
          value.currencies || { "-": null }
        ).join(",");

        infoCountries.push({ ...country });
        infoCountries.sort((a, b) => a.name.localeCompare(b.name));
      });

      const wb = new xl.Workbook();
      const ws = wb.addWorksheet("Countries");

      const numberStyle = wb.createStyle({
        numberFormat: "#,##0.00",
      });

      const titleStyle = wb.createStyle({
        font: {
          bold: true,
          color: "#4F4F4F",
          size: 16,
        },
        alignment: {
          horizontal: "center",
        },
      });

      const columnsStyle = wb.createStyle({
        font: {
          bold: true,
          color: "#808080",
          size: 12,
        },
      });

      ws.cell(1, 1, 1, 4, true).string("Countries List").style(titleStyle);

      const headingColumnNames = ["Name", "Capital", "Area", "Currencies"];

      let headingColumnIndex = 1;
      headingColumnNames.forEach((heading) => {
        ws.cell(2, headingColumnIndex++)
          .string(heading)
          .style(columnsStyle);
      });

      let rowIndex = 3;
      infoCountries.forEach((record) => {
        let columnIndex = 1;
        Object.keys(record).forEach((columnName) => {
          if (columnIndex === 3) {
            ws.cell(rowIndex, columnIndex++)
              .number(record[columnName])
              .style(numberStyle);
          } else {
            ws.cell(rowIndex, columnIndex++).string(record[columnName]);
          }
        });
        rowIndex++;
      });

      ws.column(3).setWidth(13);

      wb.write("Countries List.xlsx");
    })
    .catch((e) => console.log(e));
}

getCountries();
