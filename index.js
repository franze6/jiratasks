const ExcelJS = require('exceljs');
const fetch = (...args) => import('node-fetch').then(({default: fetch}) => fetch(...args));
const https = require('https');
const { exec } = require('child_process');

const agent = new https.Agent({
  rejectUnauthorized: false
});

const login = 'login'
const password = 'password'
// Код проекта, который можно посмотреть в веб-интерфейсе
const projectCode = 'PRCODE'
// Логин пользователя, по которому производится поиск, по-умолчанию равен логину для авторизации
const user = login

const filename = 'tasks.xlsx'

const date = new Date();
const days = new Date(date.getFullYear(), date.getMonth()+ 1, 0).getDate();
const startDate = `${date.getFullYear()}-${date.getMonth() + 1}-1`;
const endDate = `${date.getFullYear()}-${date.getMonth() + 1}-${days}`;

const base64string = Buffer.from(`${login}:${password}`).toString('base64');

const body = JSON.stringify({
  jql: `project = ${projectCode} AND issuetype = Task AND 'Start date' >= ${startDate} AND 'Start date' <= ${endDate} AND assignee in (${user}) ORDER BY priority DESC, updated DESC`,
  startAt: 0,
  maxResults: 50,
  fields: [
    "summary"
  ]
})

fetch('https://task.corp.dev.vtb/rest/api/2/search', {
  method: 'POST',
  headers: {
    Authorization: `Basic ${base64string}`,
    'Content-Type': 'application/json'
  },
  body,
  agent
}).then(res => res.json()).then(res => {
  return res.issues.map(item => ({name: item.key, title: item.fields.summary}))
}).then(res => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Tasks');
  sheet.columns = [
    { header: 'Name', key: 'name', width: 32 },
    { header: 'Title', key: 'title', width: 100 }
  ];
  res.forEach(item => {
    sheet.addRow({...item});
  })
  workbook.xlsx.writeFile(filename).then(() => {
    exec(filename)
  });
})

