const express = require('express');
const http = require('http');
const bodyParser = require('body-parser');
const multiparty = require('multiparty');
const xlsx = require('xlsx');
const fs = require('fs')
const app = express();
 
app.use(bodyParser.json());
//bodyParser.json() : 'application/json' 방식의 Content-Type 데이터를 받아준다

app.use(bodyParser.urlencoded({
    limit: '150mb',
    extended: false,
}));
// bodyParser.urlencoded({...}):  'application/x-www-form-urlencoded' 방식의 Content-Type 데이터를 받아준다.
// limit : 파일 전송시 용량을 제한하거나 풀어주기 위해 사용
// extended(false) : 내부적으로  querystring library(키 값에 깊이 모두 포함) 를 사용
// extended(true) : 내부적으로 qs library(깊이를 나눠서 정리)를 사용하여 URL-encoded data를 파싱

 
app.get('/', (req, res, next) => {
    let contents = '';
    contents += '<html><body>';
    contents += '   <form action="/" method="POST" enctype="multipart/form-data">';
    contents += '       <input type="file" name="xlsx" />';
    contents += '       <input type="submit" />';
    contents += '   </form>';
    contents += '</body></html>';
 
    res.send(contents);
});
 
app.post('/', (req, res, next) => {
    const resData = {};
 
    //multiparty를 이용해 Form 데이터를 처리
    //이 때, autoFiles를 true로 지정하면 POST 방식으로 전달된 파일만 처리하도록 할 수 있다.
        const form = new multiparty.Form({
            autoFiles: false,
            // uploadDir : 'temp/'
    });
 
    form.on('file', (name, file) => {
        // xlsx를 이용해 전달된 파일을 객체로 변환
        //객체에는 파일 정보, 테마, 시트별 데이터 등의 정보가 담김
        const workbook = xlsx.readFile(file.path);
        const sheetnames = Object.keys(workbook.Sheets);
 
        let i = sheetnames.length;
 
        while (i--) {
            const sheetname = sheetnames[i];
            resData[sheetname] = xlsx.utils.sheet_to_json(workbook.Sheets[sheetname]);
        }
    });
 
    form.on('close', () => {
        res.send(resData);
    });
 
    form.parse(req);{
        fs.writeFileSync('./test.json', Object, 'utf-8');
    }
});
 
http.createServer(app).listen(3000, () => {
    console.log('HTTP server listening on port ' + 3000);
});

module.exports = app;