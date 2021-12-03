<template>
  <div>
    <button type="button" class="btn btn-danger" @click="RunExcelJSExport">Run ExcelJS Export</button>
    <table class="table table-bordered">
      <thead>
        <tr>
          <th scope="col">Title</th>
          <th scope="col">First Name</th>
          <th scope="col">Last Name</th>
          <th scope="col">City</th>
          <th scope="col">State</th>
          <th scope="col">Position</th>
          <th scope="col">Birth Date</th>
          <th scope="col">Hire Date</th>
          <th scope="col">Sale Amount</th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="employee in employees" v-bind:key="employee.id">
          <th scope="row">{{employee.Prefix}}</th>
          <td>{{employee.FirstName}}</td>
          <td>{{employee.LastName}}</td>
          <td>{{employee.City}}</td>
          <td>{{employee.State}}</td>
          <td>{{employee.Position}}</td>
          <td>{{employee.BirthDate}}</td>
          <td>{{employee.HireDate}}</td>
          <td>{{employee.SaleAmount}}</td>
        </tr>
      </tbody>
    </table>
  </div>
</template>

<script>
// import ExcelJS from "exceljs";
// import { saveAs } from "file-saver";

// const workbook = new ExcelJS.Workbook();
// const sheet = workbook.addWorksheet('My Sheet', {
//   headerFooter:{firstHeader: "Hello Exceljs", firstFooter: "Hello World"}
// });
export default {
  name: "TableContent",
  computed: {
    employees() {
      return this.$store.getters.employees
    }
  },
  methods: {
    RunExcelJSExport() {
      // var workbook = new ExcelJS.Workbook();
      // // const worksheet = workbook.addWorksheet('Sheet 1');
      // workbook.xlsx.writeBuffer().then(function (buffer) {
      //   saveAs(
      //     new Blob([buffer], { type: "application/octet-stream" }),
      //     "test.xlsx"
      //   );
      // });
      this.$store.dispatch("RunExcelJSExport");
    },
  },
};
</script>

<style>
</style>