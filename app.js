const express = require('express');
const bodyParser = require('body-parser');
const excel = require('exceljs');

const app = express();
const port = 5000;

app.use(bodyParser.json());

const workbook = new excel.Workbook();
const filePath = './users.xlsx';
let usersData = [];

function readExcel() {
  workbook.xlsx.readFile(filePath).then(() => {
    const worksheet = workbook.getWorksheet(1);
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber !== 1) {
        const [id, name, email, number] = row.values;
        usersData.push({ id, name, email, number });
      }
    });
  });
}

function writeExcel() {
  const worksheet = workbook.getWorksheet(1);
  worksheet.spliceRows(2, worksheet.rowCount - 1);
  usersData.forEach((user) => {
    worksheet.addRow([user.id, user.name, user.email, user.number]);
  });
  return workbook.xlsx.writeFile(filePath);
}

app.post('/create', (req, res) => {
  const { name, email, number } = req.body;
  const id = usersData.length + 1;
  const newUser = { id, name, email, number };
  usersData.push(newUser);
  writeExcel().then(() => res.send(newUser));
});

app.delete('/delete/:id', (req, res) => {
  const id = parseInt(req.params.id);
  usersData = usersData.filter((user) => user.id !== id);
  writeExcel().then(() => res.send(`User with ID ${id} deleted successfully`));
});

app.put('/update/:id', (req, res) => {
  const id = parseInt(req.params.id);
  const { name, email, number } = req.body;
  const updatedUser = { id, name, email, number };
  usersData = usersData.map((user) => (user.id === id ? updatedUser : user));
  writeExcel().then(() => res.send(updatedUser));
});

app.get('/read', (req, res) => {
  readExcel();
  res.send(usersData);
});

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
