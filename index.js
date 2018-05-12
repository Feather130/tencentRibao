const request = require('request');
const ejsexcel = require('ejsexcel');
const fs = require('fs');
const exlBuf = fs.readFileSync('./template.xlsx');
const config = fs.readFileSync('./config.txt', { encoding: 'utf8' });

const indexs={
  tkIndex:config.indexOf("tk="),
  sdateIndex:config.indexOf("sdate="),
  edateIndex:config.indexOf("edate="),
  cookieIndex:config.indexOf("cookie="),
  excelIndex:config.indexOf("excel=")
}

const params = {
  id: config.slice(3, indexs.tkIndex - 1),
  tk: config.slice(indexs.tkIndex + 3, indexs.sdateIndex - 1),
  sdate: config.slice(indexs.sdateIndex + 6, indexs.edateIndex - 1),
  edate: config.slice(indexs.edateIndex + 6, indexs.cookieIndex - 1),
  cookie: config.slice(indexs.cookieIndex + 7, indexs.excelIndex - 1),
  excel: config.slice(indexs.excelIndex + 6)
};

function dailyWork(params) {
  const options = {
    url: "http://e.qq.com/ec/api.php",
    qs: {
      mod: "report",
      act: "campaign",
      owner: params.id,
      advertiser_id: params.id,
      unicode: "true",
      g_tk: params.tk,
      datetype: "1",
      format: "json",
      page: "1",
      pagesize: "100",
      sdate: params.sdate,
      edate: params.edate,
      status: "999",
      searchcname: "",
      reportonly: "0",
      time_rpt: "1"
    },
    headers: {
      "Cache-Control": "no-cache",
      Host: "e.qq.com",
      "X-Requested-With": "XMLHttpRequest",
      "User-Agent":
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Safari/537.36",
      Connection: "keep-alive",
      "Accept-Language": "zh-CN,zh;q=0.9",
      Accept: "application/json, text/javascript, */*; q=0.01",
      Cookie: params.cookie
    }
  };
  request(options, (error, response, body) => {
    if (error) throw new Error(error);
    const data = JSON.parse(body).data.list;
    const daily = [];
    data.map(item => {
      const name = item.campaignname;
      const validclickcount = item.validclickcount;
      const viewcount = item.viewcount;
      const activatedCount = item.activated_count;
      const cost = item.cost / 100 || "-";
      const clickRate =
        parseFloat((validclickcount / viewcount * 100).toFixed(2)) || "-";
      const clickPrice = parseFloat((cost / validclickcount).toFixed(2)) || "-";
      const activatedPrice =
        parseFloat((cost / activatedCount).toFixed(2)) || "-";
      const activatedRate = parseFloat(
        (activatedCount / validclickcount * 100).toFixed(2)
      );
      daily.push({
        date: params.sdate,
        name,
        viewcount, //曝光量
        validclickcount, //点击量
        clickRate, //点击率
        clickPrice, //点击均价
        activatedCount, //激活总量
        activatedRate: isNaN(activatedRate) ? "-" : activatedRate, //点击激活率
        activatedPrice: activatedPrice === Infinity ? 0 : activatedPrice, //激活均价
        cost //花费
      });
    });
    ejsexcel
      .renderExcel(exlBuf, daily)
      .then(io => {
        fs.writeFileSync(`./${params.excel}.xlsx`, io);
        console.log('日报已生成')
      })
      .catch(err => console.error(err));
  });
}
dailyWork(params);