<template>
  <div @drop="_drop" @dragenter="_suppress" @dragover="_suppress">
    <div class="row">
      <div class="col-xs-12">
        <form class="form-inline">
          <div class="form-group">
            <input type="file" class="form-control" id="file" :accept="SheetJSFT" @change="_change"/>
          </div>
        </form>
      </div>
    </div>
    <div class="row">
      <div class="col-xs-12">
        <button :disabled="data.length ? false : true" class="btn btn-success" @click="_export">Export</button>
      </div>
    </div>
    <div class="row">
      <div class="col-xs-12">
        <div class="table-responsive">
          <table class="table table-striped">
            <thead>
            <tr>
              <th v-for="c in cols" :key="c.key">{{ c.name }}</th>
            </tr>
            </thead>
            <tbody>
            <tr v-for="(r, key) in data" :key="key">
              <td class="td_table" v-for="c in cols" :key="c.key"> {{ r[c.key] }}</td>
            </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div>
</template>

<script>
import * as XLSX from "xlsx";

const make_cols = refstr => Array(XLSX.utils.decode_range(refstr).e.c + 1).fill(0).map((x, i) => ({
  name: XLSX.utils.encode_col(i),
  key: i
}));

const _SheetJSFT = [
  "xlsx", "xlsb", "xlsm", "xls", "xml", "csv", "txt", "ods", "fods", "uos", "sylk", "dif", "dbf", "prn", "qpw", "123", "wb*", "wq*", "html", "htm"
].map(function (x) {
  return "." + x;
}).join(",");

export default {
  data() {
    return {
      data: ["SheetJS".split(""), "1234567".split("")],
      cols: [
        {name: "A", key: 0},
        {name: "B", key: 1},
        {name: "C", key: 2},
        {name: "D", key: 3},
        {name: "E", key: 4},
        {name: "F", key: 5},
        {name: "G", key: 6},
      ],
      SheetJSFT: _SheetJSFT
    };
  },
  methods: {
    _suppress(evt) {
      evt.stopPropagation();
      evt.preventDefault();
    },
    _drop(evt) {
      evt.stopPropagation();
      evt.preventDefault();
      const files = evt.dataTransfer.files;
      if (files && files[0]) this._file(files[0]);
    },
    _change(evt) {
      const files = evt.target.files;
      if (files && files[0]) this._file(files[0]);
    },
    _export(evt) {

      /* преобразовать состояние в книгу */
      const ws = XLSX.utils.aoa_to_sheet(this.data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "SheetJS");

      /* сгенерировать файл и отправить клиенту */
      XLSX.writeFile(wb, "sheetjs.xlsx");
    },
    _file(file) {

      /* Шаблон для настройки FileReader */
      const reader = new FileReader();
      reader.onload = (e) => {

        /* Анализировать данные */
        const bstr = e.target.result;
        const wb = XLSX.read(bstr, {type: 'binary'});

        /* Получить первый рабочий лист */
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];

        /* Преобразование массива массивов */
        const data = XLSX.utils.sheet_to_json(ws, {header: 1});

        /* Состояние обновления */
        this.data = data;
        this.cols = make_cols(ws['!ref']);
      };
      reader.readAsBinaryString(file);
    }
  }
};
</script>

<style>
.table, .td_table{
  border: solid 1px black;
  padding: 3px;
}
</style>
