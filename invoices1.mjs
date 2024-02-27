import fetch from 'node-fetch'
import fs from 'fs'
import moment from 'moment'
import puppeteer from 'puppeteer'
import xlsx, { read, utils } from 'xlsx'
import cheerio from 'cheerio'
import converter from 'json-2-csv'
import PDFLib from 'pdf-lib'
import { JSDOM } from 'jsdom'
import util from 'util'
//"C:\Users\mis1.ryd\OneDrive - Saudi German Hospital (1)\invoices1.mjs"
const writeFile = util.promisify(fs.writeFile)
Date.prototype.addDays = function (days) {
  var date = new Date(this.valueOf());
  date.setDate(date.getDate() + days);
  return date;
}

let sbsmapping=JSON.parse(fs.readFileSync("C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Maps/nphies CodeSet/SBSMapping.json"))
let prctmissing= JSON.parse(fs.readFileSync("C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Maps/nphies CodeSet/PRCTMapping.json"))

const bl = ['FMLAB-2030',
  'FMLAB-2029',
  'FMLAB-2023',
  'FMLAB-2012',
  'FMLAB-1023',
  'FMLAB-2073',
  'FMLAB-5023',
  'FMLAB-2056',
  'FMLAB-2060',
  'FMLAB-2015',
  'FMLAB-2028',
  'FMLAB-2065',
  'FMLAB-1065',
  'FMLAB-1022',
  'FMLAB-2064',
  'FMLAB-1054',
  'FMLAB-2005',
  'FMLAB-2100', 'FMPOCT-2'
]


function getUniqueByBillNo(inputArray) {
return inputArray.reduce((unique,item)=>{
const exists=unique.some((u)=>u.BillNo===item.BillNo)
if(!exists){
  unique.push(item)
}
return unique
},[])
}


function convertJSONtoCSV(jsonData) {
  const csvRows = []

  // Extract header row
  const headers = Object.keys(jsonData[0])
  csvRows.push(headers.join(','))

  // Convert each object to CSV row
  jsonData.forEach((obj) => {
    const values = headers.map((header) => {
      const escapedValue = String(obj[header]).replace(/"/g, '\\"')
      return `"${escapedValue}"`
    })
    csvRows.push(values.join(','))
  })

  // Convert CSV rows to a string
  const csvString = csvRows.join('\n')

  return csvString
}

async function mergePDFDocuments(documents) {
  // const PDFDocument= PDFDocument
  const mergedPdf = await PDFLib.PDFDocument.create()
  let t = 0
  for (let document of documents) {
    t++
    //console.log("mergeing",t,documents.length)
    document = await PDFLib.PDFDocument.load(document)

    const copiedPages = await mergedPdf.copyPages(document, document.getPageIndices())
    for (let page of copiedPages) {
      mergedPdf.addPage(page)
    }
    //copiedPages.forEach((page) => mergedPdf.addPage(page))
  }

  return await mergedPdf.save()
}

function numberbetween(x) {
  const mySubString = x.substring(
    x.indexOf('(') + 1,
    x.lastIndexOf(')')
  )
  return mySubString == '' ? 0 : +mySubString
}

function compareNumbers(a, b) {
  return numberbetween(b) - numberbetween(a)
}
const options = {
  prependHeader: true
  // expandArrayObjects:true
}
//const logos=JSON.parse(fs.readFileSync('./logos.json'))

const branchs = [

  //    { Link: 'http://130.3.2.208', Name: 'Aseer', Password: 'EXPIRED', User: '20214115', type: 'HCP', isnph: true, hasMR: true },
  //    { Link: 'http://130.2.10.21', Name: 'Riyadh', Password: 'EXPIRED', User: '20214115', type: 'HCP', isnph: true, hasMR: true },
  //    { Link: 'http://130.4.1.16', Name: 'Madinah', Password: 'EXPIRED', User: '20214115', type: 'HCP', isnph: true, hasMR: true  },
  //     { Link: 'http://130.8.2.18', Name: 'Hail', Password: 'Password1', User: '20214115', type: 'HCP', isnph: true },
  // { Link: 'http://130.1.2.75', Name: 'Hai', Password: '20210826E', User: '20210826', type: 'HCP', isnph: true  },
  { Link: 'http://10.14.5.51', Name: 'Makkah', Password: '123456789P', User: '20210674', type: 'HCP'  },
  //   { Link: 'http://130.1.2.27', Name: 'Jeddah', Password: '2', User: '20214115', type: 'HCP', hasMR: true , isnph: true },
  // { Link: 'http://hisabha.sghgroup.com', Name: 'Abha', Password: 'ALSH50301', User: '20214849'  },
  // { Link: 'http://130.1.2.153', Name: 'Beverly', Password: 'ZKHASSAN', User: '20210674'  }

]
function ExcelDateToJSDate(serial) {
  const utc_days = Math.floor(serial - 25569)
  const utc_value = utc_days * 86400
  const date_info = new Date(utc_value * 1000)

  const fractional_day = serial - Math.floor(serial) + 0.0000001

  let total_seconds = Math.floor(86400 * fractional_day)

  const seconds = total_seconds % 60

  total_seconds -= seconds

  const hours = Math.floor(total_seconds / (60 * 60))
  const minutes = Math.floor(total_seconds / 60) % 60

  return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds)
}
function getlastday(month) {
  let startofmonth = new Date('2023' + '-' + month + '-2')
  let lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 1)
  startofmonth = startofmonth.toISOString().split('T')[0]
  lastDayOfMonth = lastDayOfMonth.toISOString().split('T')[0]
  const dd = lastDayOfMonth.slice(lastDayOfMonth.lastIndexOf('-') + 1)
  return dd
}
function datesInPeriod(start, end) {
  const startDate = moment(start, 'DD-MMM-YYYY')
  // console.log(startDate)
  const endDate = moment(end, 'DD-MMM-YYYY')

  // Initialize empty array to store dates
  const datesArray = []

  // Loop through dates and push to array
  while (startDate <= endDate && startDate <= new moment().add(-1, 'days')) {
    datesArray.push(startDate.format('DD-MMM-YYYY'))
    startDate.add(1, 'days')
    // console.log(datesArray)
  }

  return datesArray
}

function compareAge(a, b) {
  return a.Company - b.Company
}
function getMonthStartAndEndDates(startDate, endDate) {
  const start = new Date(startDate)
  const end = new Date(endDate)
  const result = []

  // Set the start date to the beginning of the month
  start.setDate(1)

  // Loop through each month
  while (start <= end) {
    const month = start.getMonth()
    const year = start.getFullYear()

    // Calculate the last day of the current month
    const lastDay = new Date(year, month + 1, 0).getDate()

    // Create a new object with the start and end dates of the current month
    const startDateFormatted = formatDate(start)
    const endDateFormatted = formatDate(new Date(year, month, lastDay))

    result.push({
      start: startDateFormatted,
      end: endDateFormatted
    })

    // Move to the next month
    start.setMonth(month + 1)
  }

  return result
}

// Helper function to format date as MM/DD/YYYY
function formatDate(date) {
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  const year = date.getFullYear()

  return `${month}/${day}/${year}`
}
function getMonthStartAndEndDates2(startDate, endDate) {
  const start = new Date(startDate)
  const end = new Date(endDate)
  const result = []

  // Set the start date to the beginning of the month
  start.setDate(1)

  // Loop through each month
  while (start <= end) {
    const month = start.getMonth()
    const year = start.getFullYear()

    // Calculate the last day of the current month
    const lastDay = new Date(year, month + 1, 0).getDate()

    // Create a new object with the start and end dates of the current month
    const startDateFormatted = formatDate2(start)
    const endDateFormatted = formatDate2(new Date(year, month, lastDay))

    result.push({
      start: startDateFormatted,
      end: endDateFormatted
    })

    // Move to the next month
    start.setMonth(month + 1)
  }

  return result
}

// Helper function to format date as MM/DD/YYYY
function formatDate2(date) {
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  const year = date.getFullYear()

  return `${year}-${month}-${day}`
}
const delay = ms => new Promise(res => setTimeout(res, ms))
const browser = await puppeteer.launch({ headless: false })
const promises = []
branchs.forEach(async branch => {
  const loc = branch.Link
  const br = branchs.filter(a => a.Link == loc)[0].Name
  const code = branchs.filter(a => a.Link == loc)[0].s
  const pass = branchs.filter(a => a.Link == loc)[0].Password
  const user = branchs.filter(a => a.Link == loc)[0].User
  const type = branchs.filter(a => a.Link == loc)[0].type
  const page = await browser.newPage()
  await page.setDefaultNavigationTimeout(0)
  await page.goto(loc + '/HISLOGIN/Home/LogOff')
  await page.reload()
  let cookie = await page.cookies()
  // console.log(br,cookie,cookie.length,cookie[0].value)
  if (cookie.length <= 1 || cookie[0].value == 1) {
    // console.log(loc,user,pass,page)
    await page.type('#txtusername', user)
    await page.type('#txtpassword', pass)
    await page.click('#btn-login')
    // await page.screenshot({ path: 'example.png' })
    // console.log(page.url())
    await delay(30000)
    await page.reload()

    cookie = await page.cookies()
    // await page.close()
  }
  async function tts() {
    await page.reload()

    cookie = await page.cookies()
    console.log(br, "cookie reloaded")
  }
  // const browser= await puppeteer.launch()
  setInterval(() => {
    tts


  }, 10 * 60 * 1000);
  async function refreshAndGetCookie(loc, user, pass) {
    const page = await browser.newPage()

    await page.goto(loc + '/HISLOGIN/Home/LogOff')
    // await page.reload()
    let cookie = await page.cookies()
    // console.log(br,cookie,cookie.length,cookie[0].value)
    if (cookie.length <= 1 || cookie[0].value == 1) {
      console.log(loc, user, pass, page)
      await page.type('#txtusername', user)
      await page.type('#txtpassword', pass)
      await page.click('#btn-login')
      await page.screenshot({ path: 'example.png' })
      // console.log(page.url())
      await page.reload()
      // await page.waitForNavigation( );

      cookie = await page.cookies()
    }// console.log(cookie); // or do whatever you want with the new cookie
    // console.log(br,cookie)
    return cookie
  }

  function flattenObject(obj, prefix = '') {
    return Object.keys(obj).reduce((acc, key) => {
      const propKey = prefix ? `${prefix}.${key}` : key
      if (typeof obj[key] === 'object' && !Array.isArray(obj[key])) {
        Object.assign(acc, flattenObject(obj[key], propKey))
      } else {
        acc[propKey] = obj[key]
      }
      return acc
    }, {})
  }

  async function ptDetails(pin) {
    const x = await fetch(loc + '/HISFRONTOFFICE/Registration/GetPatientDetails', {
      headers: {
        accept: 'application/json, text/javascript, */*; q=0.01',
        'accept-language': 'en-US,en;q=0.9,ar;q=0.8',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'x-requested-with': 'XMLHttpRequest'
      },
      referrer: 'http://130.2.10.21/HISFRONTOFFICE/Registration',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: 'RegistrationNo=' + pin,
      method: 'POST',
      mode: 'cors',
      credentials: 'include'
    }).then(a => a.json()).then(a => a.ReturnData)
    return x
    // x.PTDetails.ResidenceId
  }

  async function cashlist() {
    let comps = await fetch(loc + '/HISARADMIN/Common/get_common_list', {
      headers: {
        accept: 'application/json, text/javascript, */*; q=0.01',
        'accept-language': 'en-US,en;q=0.9,ar;q=0.8',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
      },
      referrer: 'http://130.3.2.208/HISARADMIN/CompanyProfile',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: 'id=0&ctype=-201',
      method: 'POST',
      mode: 'cors',
      credentials: 'include'
    }).then(a => a.json()).then(a => a.CL)
    comps = comps.find(a => a.Name.toLowerCase().includes('cash')).Id

    let x = await fetch(loc + '/HISARADMIN/CompanyProfile/GetCompanyMasterList', {
      headers: {
        accept: 'application/json, text/javascript, */*; q=0.01',
        'accept-language': 'en-US,en;q=0.9,ar;q=0.8',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
      },
      referrer: loc + '/HISARADMIN/CompanyProfile',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: 'cid=' + comps,
      method: 'POST',
      mode: 'cors',
      credentials: 'include'
    }).then(a => a.json()).then(a => a.Res)
    x = x.map(a => a.Name)

    return { list: x, comps }
  }

  async function ptDetails(pin) {
    const x = await fetch(loc + '/HISFRONTOFFICE/Registration/GetPatientDetails', {
      headers: {
        accept: 'application/json, text/javascript, */*; q=0.01',
        'accept-language': 'en-US,en;q=0.9,ar;q=0.8',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
      },
      referrer: 'http://130.2.10.21/HISFRONTOFFICE/Registration',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: 'RegistrationNo=' + pin,
      method: 'POST',
      mode: 'cors',
      credentials: 'include'
    }).then(a => a.json()).then(a => a.ReturnData)
    return x
    // x.PTDetails.ResidenceId
  }

  async function getbillexcel(billNo) {
    const res = await fetch(loc + '/HISARADMIN/Reports/Reports.aspx', {
      headers: {
        accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'accept-language': 'en-US,en;q=0.9',
        'cache-control': 'max-age=0',
        'content-type': 'application/x-www-form-urlencoded',
        'upgrade-insecure-requests': '1',
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
      },
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: 'billno=' + billNo + '&rtype=' + 21 + '&ismain=1',
      method: 'POST',
      mode: 'cors',
      credentials: 'include'
    }).then(a => a.text())
    const ctrl = res.slice(
      res.indexOf('ControlID=') + 'ControlID='.length,
      res.indexOf('&Mode')
    ).substring(0, 32)
    const xx = await fetch(loc + '/HISARADMIN/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + ctrl + '&Mode=true&OpType=Export&FileName=IPInvoiceMainNPD&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML', {
      headers: {
        accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'accept-language': 'en-US,en;q=0.9',
        'cache-control': 'no-cache',
        pragma: 'no-cache',
        'upgrade-insecure-requests': '1',
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
      },
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: null,
      method: 'GET',
      mode: 'cors',
      credentials: 'include'
    }).then(a => a.arrayBuffer())
    return xx
  }

  async function getbillno(PIN) {
    let z = []
    const x =
      await fetch(loc + '/HISIPBILL/IPBILL/CashToCompany/GetAdmitDate?id=' + PIN + '&_=1683613288392', {
        headers: {
          accept: '*/*',
          'accept-language': 'en-US,en;q=0.9',
          'x-requested-with': 'XMLHttpRequest',
          cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
        },
        referrer: loc + 'HISIPBILL/IPBILL/CashToCompany/',
        referrerPolicy: 'strict-origin-when-cross-origin',
        body: null,
        method: 'GET',
        mode: 'cors',
        credentials: 'include'
      }).then(a => a.json())
    z.push(x)
    const y = await fetch(loc + '/HISIPBILL/IPBILL/CompanyToCash/GetAdmitDate?id=' + PIN + '&_=1683613288392', {
      headers: {
        accept: '*/*',
        'accept-language': 'en-US,en;q=0.9',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
      },
      referrer: loc + 'HISIPBILL/IPBILL/CompanyToCash/',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: null,
      method: 'GET',
      mode: 'cors',
      credentials: 'include'
    }).then(a => a.json())
    z.push(y)
    z = z.flat()
    return z
    // filter DischargeDateTime to get billno
  }
  async function get_IPBill_Cats(billNo) {
    return await fetch(loc + '/HISARADMIN/ARIPBillCorrection/get_IPBill_Services', {
      headers: {
        accept: 'application/json, text/javascript, */*; q=0.01',
        'accept-language': 'en-US,en;q=0.9',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
      },
      referrer: loc + 'HISARADMIN/ARIPBillCorrection',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: 'billno=' + billNo,
      method: 'POST',
      mode: 'cors',
      credentials: 'include'
    }).then(a => a.json()).then(a => a.Res)
  }
  async function getIPBillServices(billno, cat) {
    return await fetch(loc + '/HISARADMIN/ARIPBillCorrection/get_IPBill_Items', {
      headers: {
        accept: 'application/json, text/javascript, */*; q=0.01',
        'accept-language': 'en-US,en;q=0.9',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
      },
      referrer: loc + 'HISARADMIN/ARIPBillCorrection',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: 'billno=' + billno + '&serid=' + cat,
      method: 'POST',
      mode: 'cors',
      credentials: 'include'
    }).then(a => a.json()).then(a => a.Res)
  }

  async function ipbillF(visit) {
    // ViewBillsRep
    // "http://130.4.1.16/HISIPBILL/IPBILL/IPBilling/ViewBills?typ=IN&ipid=148833&sdate1=27+Feb+2023+23%3A55%3A53&sdate2=3%2F2%2F2023+8%3A43%3A13+AM&comp=24554&grade=54582&tar=23&catid=23&pckid=0&_=1677735780091"
    const ipbill = await fetch(loc + '/HISIPBILL/IPBILL/IPBilling/ViewBillsAR?typ=' + visit.PatientType + '&ipid=' + visit.id + '&sdate1=' + visit.AdmitDateTime + '&sdate2=' + visit.DischargeDateTime + '&comp=' + visit.CompanyID + '&grade=' + visit.GradeID + '&tar=' + visit.TariffID + '&catid=' + visit.CategoryID + '&pckid=&_=1675519130840', {
      headers: {
        accept: '*/*',
        'accept-language': 'en-US,en;q=0.9',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
      },
      referrer: loc + '/HISIPBILL/IPBILL/Reports/ViewInpatientBillReport',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: null,
      method: 'GET',
      mode: 'cors',
      credentials: 'include'
    }).then((response) => response.text())
    return ipbill
    // writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/IPIDS/'+br+'_' +visit.id+"_ipid.html",ipbill)
  }

  async function checkDate(PIN) {
    try {
      const pin = +PIN + 0 == PIN ? PIN : PIN.split('.').pop()
      const xxx = await fetch(loc + '/HISIPBILL/IPBILL/IPBilling/ViewDates?id=' + pin + '&_=1675519487359', {
        headers: {
          accept: '*/*',
          'accept-language': 'en-US,en;q=0.9',
          'x-requested-with': 'XMLHttpRequest',
          cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
        },
        referrer: loc + '/HISIPBILL/IPBILL/Reports/ViewInpatientBillReport',
        referrerPolicy: 'strict-origin-when-cross-origin',
        body: null,
        method: 'GET',
        mode: 'cors',
        credentials: 'include'
      }).then((a) => a.json())
      // .catch(a=>a)
      return xxx
    } catch (error) {
      console.log('failed to get date for ' + PIN)
    }
    // .find(a=>a.id==IPID)
  }
  // let cookie=await refreshAndGetCookie(loc,user,pass)
  const comps = await fetch(loc + '/HISARADMIN/Common/get_common_list', {
    headers: {
      accept: 'application/json, text/javascript, */*; q=0.01',
      'accept-language': 'en-US,en;q=0.9',
      'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
      'x-requested-with': 'XMLHttpRequest',
      cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
    },
    referrer: loc + '/HISARADMIN/CategoryWiseExtraction',
    referrerPolicy: 'strict-origin-when-cross-origin',
    body: 'id=0&ctype=-200',
    method: 'POST',
    mode: 'cors',
    credentials: 'include'
  }).then(res => res.json())

  writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Eclaim/2023/eclaim/' + br + '.json', JSON.stringify(comps))

  const Depts = await fetch(loc + '/HISMCRS/ManagementReports/ARReports/DoctorRevenueOP', {
    headers: {
      accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
      'accept-language': 'en-US,en;q=0.9,ar;q=0.8',
      'cache-control': 'max-age=0',
      'upgrade-insecure-requests': '1',
      cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
    },
    referrerPolicy: 'strict-origin-when-cross-origin',
    body: null,
    method: 'GET',
    mode: 'cors',
    credentials: 'include'
  }).then(a => a.text())

  // .then(a=>Array.from(a.querySelector("#DepartmentId").childNodes).map(a=>({val:a.value,Name:a.innerText})))
  writeFile(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Depts/${br}_Depts.json`, Depts.slice(Depts.indexOf('"DepartmentList"') + '"DepartmentList"'.length + 1, Depts.indexOf('}]}')) + '}]')

  // all drs
  // downloadBlob(JSON.stringify(Depts),br+"_"+'Depts.json','text/html')
  const doctors = await fetch(loc + '/HISDM/DM/Generic/DoctorList?_=1681285604651', {
    headers: {
      accept: '*/*',
      'accept-language': 'en-US,en;q=0.9,ar;q=0.8',
      'x-requested-with': 'XMLHttpRequest',
      cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
    },
    referrer: loc + '/HISDM/DM/Setup/ResidentDoctor',
    referrerPolicy: 'strict-origin-when-cross-origin',
    body: null,
    method: 'GET',
    mode: 'cors',
    credentials: 'include'
  }).then(a => a.json())
  writeFile(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Depts/${br}_Doctors.json`, JSON.stringify(doctors))
  const emps = await fetch(loc + "/HISHR/PAYROLL/Employee/DRegularFilter?Top=500000&ID=-1&CategoryId=-1&DepartmentId=-1&PositionId=-1&NationalityId=-1&IsActive=0&_=1688546476345", {
    "headers": {
      "accept": "*/*",
      "accept-language": "en-US,en;q=0.9,ar;q=0.8",
      "content-type": "application/json; charset=utf-8",
      "x-requested-with": "XMLHttpRequest",
      cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
    },
    "referrer": "http://130.2.10.21/HISLOGIN/",
    "referrerPolicy": "strict-origin-when-cross-origin",
    "body": null,
    "method": "GET",
    "mode": "cors",
    "credentials": "include"
  }).then(a => a.json())
  writeFile(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Depts/${br}_Emps.json`, JSON.stringify(emps))
  let allcates = await fetch(loc + "/HISARADMIN/CategoryProfile/GetCategoryMasterList", {
    "headers": {
      "accept": "application/json, text/javascript, */*; q=0.01",
      "accept-language": "en-US,en;q=0.9,ar;q=0.8",
      "x-requested-with": "XMLHttpRequest",
      cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
    },
    "referrer": "http://130.2.10.21/HISARADMIN/CategoryProfile",
    "referrerPolicy": "strict-origin-when-cross-origin",
    "body": null,
    "method": "POST",
    "mode": "cors",
    "credentials": "include"
  }).then(a => a.json()).then(a => a.Res)
  let allcomps = await fetch(loc + "/HISARADMIN/CompanyProfile/GetCompanyMasterList", {
    "headers": {
      "accept": "application/json, text/javascript, */*; q=0.01",
      "accept-language": "en-US,en;q=0.9,ar;q=0.8",
      "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
      "x-requested-with": "XMLHttpRequest",
      cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
    },
    "referrer": "http://130.2.10.21/HISARADMIN/CompanyProfile",
    "referrerPolicy": "strict-origin-when-cross-origin",
    "body": "cid=0",
    "method": "POST",
    "mode": "cors",
    "credentials": "include"
  }).then(a => a.json()).then(a => a.Res)
  allcomps.forEach(a => {

    a.header = allcates.find(x => x.Id == a.CategoryID)

  })

  writeFile(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Maps/${br}_allcomps.json`, JSON.stringify(allcomps))

///////////yearrr
  const dm = br == 'Jeddah' ? 'HISDM' : 'HISDM4'
  const years = [2024]//, 2022,2021,2020,2019,2018,2017,2016,2015,2014,2013,2012,2011,2010,2009,2008,2007,2006,2005,2004,2003,2002,2001,2000,1999,1998,1997]
  //, 2022,2021]
  const months = ["Feb"].reverse()// "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" ]
  //, "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]//, "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
  // ["Jan","Feb","Mar","Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
  async function opclaims() {
    let tt=0
    let errs=[]
    for (const year of years) {
    for (const month of months) {
    for (const xx of [1, 0]) {
      for (const comp of comps.CL) {
        
          
            try {
              tt++
                if (tt % 100 == 0 && tt > 99) {
                  await page.reload()
      
                  cookie = await page.cookies()
                }
              
            
            const startofmonth = new Date('01 ' + month + ' 2022')
            const lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 0)
            const dst = lastDayOfMonth.toISOString().split('T')[0]
            const lstdd = +dst.substr(dst.length - 2) + 1

           // console.log('Downloading op eclaim ' + month + ' for comp ' + comps.CL.indexOf(comp) + ' of ' + comps.CL.length)

            let op = null
            while (!op || op.status !== 200) {
              op = await fetch(loc + '/HISARADMIN/CategoryWiseExtraction/Start_ExportToExcel', {
                headers: {
                  accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                  'accept-language': 'en-US,en;q=0.9',
                  'cache-control': 'max-age=0',
                  'content-type': 'application/x-www-form-urlencoded',
                  'upgrade-insecure-requests': '1',
                  cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
                },
                referrer: loc + '/HISARADMIN/CategoryWiseExtraction',
                referrerPolicy: 'strict-origin-when-cross-origin',
                body: `catid=${comp.Id}&comid=0&type=3&aftercl=${xx}&fdate=01-${month}-${year}&tdate=${lstdd}-${month}-${year}`,
                method: 'POST',
                mode: 'cors',
                credentials: 'include'
              })
            }
            const op2 = await op.arrayBuffer()// (res1 => res1.arrayBuffer()).catch(a=>console.log(comp.Id + "&comid=0&type=3&aftercl=" + xx + "" + "&fdate=01-" + month + "-" + year + "&tdate=" + lstdd + "-" + month + "-" + year))

            // downloadBlob(op2, `${br}_${month}-${year}_${comp.Id}_${xx}_OP.xlsx`, 'application/excel')
            const optod = Buffer.from(op2)
            let sheetlenght=xlsx.utils.sheet_to_json( xlsx.readFile(op2).Sheets.Sheet1).length
            if (sheetlenght&&sheetlenght>0
              //optod.byteLength > 2832
              ) {
                fs.writeFileSync(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Eclaim/2023/eclaim/${br}_${month}-${year}_${comp.Id}_${xx}_OP.xlsx`, optod)
              console.log(`${br} Downloading OP eclaim ${month} for comp ${comps.CL.indexOf(comp)} of ${comps.CL.length} not empty`)
            } else {
             // console.log(`${br} ignoring OP eclaim ${month} for comp ${comps.CL.indexOf(comp)} of ${comps.CL.length} is empty`)
            }
            } catch (error) {
              errs.push(error)
              console.log(error)

              
            }
            

          }
        }
        fs.writeFileSync(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Eclaim/2023/eclaim/${br}_${month}_${year}_OPDerrors.txt`, errs.join('\n'))

      }
    }
    
  }

  async function ipclaims() {
    let errs=[]
    let tt=0
    console.log(`expected files ${comps.CL.length * months.length}`)
    for (const year of years) {
      for (const month of months) {
    for (const xx of [0, 1]) {
      for (const comp of comps.CL) {
        
            try {
              tt++
              if (tt % 100 == 0 && tt > 99) {
                await page.reload()
    
                cookie = await page.cookies()
              }
              
            
            const startofmonth = new Date(`01 ${month} 2022`)
            const lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 0)
            const dst = lastDayOfMonth.toISOString().split('T')[0]
            const lstdd = +dst.substr(dst.length - 2) + 1

            let ip = null

            while (!ip || ip.status !== 200) {
              ip = await fetch(loc + '/HISARADMIN/CategoryWiseExtraction/Start_ExportToExcelIP', {
                headers: {
                  accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                  'accept-language': 'en-US,en;q=0.9',
                  'cache-control': 'max-age=0',
                  'content-type': 'application/x-www-form-urlencoded',
                  'upgrade-insecure-requests': '1',
                  cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
                },
                referrer: loc + '/HISARADMIN/CategoryWiseExtraction',
                referrerPolicy: 'strict-origin-when-cross-origin',
                body: `catid=${comp.Id}&comid=0&type=4&aftercl=${xx}&fdate=01-${month}-${year}&tdate=${lstdd}-${month}-${year}`,
                method: 'POST',
                mode: 'cors',
                credentials: 'include'
              })
            }
            const ip2 = await ip.arrayBuffer()// (res1 => res1.arrayBuffer()).catch(a=>console.log(comp.Id + "&comid=0&type=3&aftercl=" + xx + "" + "&fdate=01-" + month + "-" + year + "&tdate=" + lstdd + "-" + month + "-" + year))
            // console.log(Buffer.from(ip2).byteLength)
            // downloadBlob(ip2, `${br}_${month}-${year}_${comp.Id}_${xx}_IP.xlsx`, 'application/excel')
            const iptod = Buffer.from(ip2)
            let sheetlenght=xlsx.utils.sheet_to_json( xlsx.readFile(ip2).Sheets.Sheet1).length
             
            if (sheetlenght&&sheetlenght>0
              
              //iptod.byteLength > 2801
              ) {
              fs.writeFileSync(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Eclaim/2023/eclaim/${br}_${month}-${year}_${comp.Id}_${xx}_IP.xlsx`, iptod)
              console.log(`${br} Downloading IP eclaim ${month} for comp ${comps.CL.indexOf(comp)} of ${comps.CL.length} not empty`)
            } else {
             // console.log(`${br} ignoring IP eclaim ${month} for comp ${comps.CL.indexOf(comp)} of ${comps.CL.length} is empty`)
            }
          } catch (error) {
              errs.push(`${br}_${month}-${year}_${comp.Id}_${xx}`)
              console.log(error)
          }
        }
        }
      
        fs.writeFileSync(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Eclaim/2023/eclaim/${br}_${month}_${year}_IPDerrors.txt`, errs.join('\n'))

      }
    }
     
  }

  async function summary() {
    for (const year of years) {
      for (const month of months) {
        console.log(`${br} Downloading summary ${month} ${year}`)
       let summ=await fetch(loc+"/HISARADMIN/CLLocking/get_CLList", {
          "headers": {
            "accept": "application/json, text/javascript, */*; q=0.01",
            "accept-language": "en-US,en;q=0.9,ar;q=0.8",
            "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
            "x-requested-with": "XMLHttpRequest",
            cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
          },
          "referrer": "http://130.2.10.21/HISARADMIN/CLLocking",
          "referrerPolicy": "strict-origin-when-cross-origin",
          "body": "cldate=01-"+month+"-"+year,
          "method": "POST",
          "mode": "cors",
          "credentials": "include"
        }).then(a=>a.json())

        fs.writeFileSync(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Summ/${br}_${month}-${year}_Summ.json`, JSON.stringify(summ))
        console.log('Done')
      }
    }
  }
  async function CashIP() {
    for (const month of months) {
      for (const year of years) {
        const startofmonth = new Date('01 ' + month + ' 2022')
        const lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 0)
        const dst = lastDayOfMonth.toISOString().split('T')[0]
        const lstdd = +dst.substr(dst.length - 2) + 1
        // let periods=[{from:1,to:2},{from:3,to:4},{from:6,to:6},{from:9,to:8},{from:11,to:10},{from:13,to:lstdd}]
        for (let d = 1; d <= lstdd; d++) {
          console.log(d, month, year)
          // let DTF=encodeURIComponent( "01 "+month+" "+year)
          // let DTT=encodeURIComponent(getlastday(month)+" "+month+" "+year)

          await fetch(loc + '/HISMCRS/ManagementReports/AuditReports/IPChargeBilledReport', {
            headers: {
              accept: '*/*',
              'accept-language': 'en-US,en;q=0.9,ar;q=0.8',
              'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
              'x-requested-with': 'XMLHttpRequest',
              cookie: cookie.map(a => a.name + '=' + a.value).join('; ')

            },
            referrer: loc + '/HISMCRS/ManagementReports/AuditReports/IPChargeBilledReport',
            referrerPolicy: 'strict-origin-when-cross-origin',
            body: 'StartDate=' + d + '-' + month + '-' + year + '&EndDate=' + d + '-' + month + '-' + year + '&ChargedType=1&ChargedORBilled=0&DoctorId=0&ServiceId=0&AccountType=1&CategoryId=0&X-Requested-With=XMLHttpRequest',
            method: 'POST',
            mode: 'cors',
            credentials: 'include'
          })
          console.log('fetched', d, month, year)
          let controlID = await fetch(loc + '/HISMCRS/ReportViewerWebForm.aspx', {
            headers: {
              accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
              'accept-language': 'en-US,en;q=0.9,ar;q=0.8',
              'upgrade-insecure-requests': '1',
              cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
            },
            referrer: loc + '/HISMCRS/ManagementReports/AuditReports/IPChargeBilledReport',
            referrerPolicy: 'strict-origin-when-cross-origin',
            body: null,
            method: 'GET',
            mode: 'cors',
            credentials: 'include'
          }).then(a => a.text()).then(
            res => res.slice(
              res.indexOf('ControlID=') + 'ControlID='.length,
              res.indexOf('&Mode')
            ).substring(0, 32))
          console.log(controlID, d, month, year)
          const link = loc + '/HISMCRS/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + controlID + '&Mode=true&OpType=Export&FileName=Report_IPCharged&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML'
          const pdf = await fetch(link
            , {
              headers: {
                accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                'accept-language': 'en-US,en;q=0.9',
                'cache-control': 'no-cache',
                pragma: 'no-cache',
                'upgrade-insecure-requests': '1',
                cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
              },
              referrerPolicy: 'strict-origin-when-cross-origin',
              body: null,
              method: 'GET',
              mode: 'cors',
              credentials: 'include'
            }).then(res => res.arrayBuffer())
          // pdfs.push(pdf)

          writeFile(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Cash/${br}_${d}_${month}-${year}_Cash_IP.xlsx`, Buffer.from(pdf))

          controlID = null
        }
      }
    }
  }
  async function opdcash() {
    for (const month of months) {
      for (const year of years) {
        console.log(br, month, year, 'opdcash')

        const startofmonth = new Date('01 ' + month + ' 2022')
        const lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 0)
        const dst = lastDayOfMonth.toISOString().split('T')[0]
        const lstdd = +dst.substr(dst.length - 2) + 1

        let opcharges = null
        while (!opcharges || opcharges.status !== 200) {
          opcharges = await fetch(loc + '/HISMCRS/ManagementReports/FinanceReports/OPRevenue2018', {
            headers: {
              accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
              'accept-language': 'en-US,en;q=0.9',
              'cache-control': 'max-age=0',
              'content-type': 'application/x-www-form-urlencoded',
              'upgrade-insecure-requests': '1',
              cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
            },
            referrer: loc + '/HISMCRS/ManagementReports/FinanceReports/OPRevenue',
            referrerPolicy: 'strict-origin-when-cross-origin',
            body: 'StartDate2=01-' + month + '-' + year + '&EndDate2=' + lstdd + '-' + month + '-' + year + '&PatientBillType2=' + 1,
            method: 'POST',
            mode: 'cors',
            credentials: 'include'
          })
        }
        const op2 = await opcharges.arrayBuffer()

        // downloadBlob(op2,br+"_"+month+"-"+year+"_"+type+"_OPRev.xlsx","application/Excel")
        fs.writeFileSync(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Cash/${br}_${month}-${year}_Cash_OP.xlsx`, Buffer.from(op2))
        // opcharges = null
      }
    }
  }

  async function getdisSheet(year, month) {
    console.log(month, year, 'Dissheet')

    const DTF = encodeURIComponent('01 ' + month + ' ' + year)
    const DTT = encodeURIComponent(getlastday(month) + ' ' + month + ' ' + year)

    const controlID = await fetch(loc + '/HISIPBILL/Areas/IPBILL/Report/DischargeReport.aspx?DTF=' + DTF + '&DTT=' + DTT + '&CAT=0&COMP=0', {
      headers: {
        accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'accept-language': 'en-US,en;q=0.9,ar;q=0.8',
        'upgrade-insecure-requests': '1',
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
      },
      referrer: 'http://130.1.2.27/HISIPBILL/IPBILL/Reports/ViewDischargeReport',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: null,
      method: 'GET',
      mode: 'cors',
      credentials: 'include'
    }).then(a => a.text()).then(
      res => res.slice(
        res.indexOf('ControlID=') + 'ControlID='.length,
        res.indexOf('&Mode')
      ).substring(0, 32))
    const link = loc + '/HISIPBILL/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + controlID + '&Mode=true&OpType=Export&FileName=DischargeReport&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML'

    const pdf = await fetch(link
      , {
        headers: {
          accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
          'accept-language': 'en-US,en;q=0.9',
          'cache-control': 'no-cache',
          pragma: 'no-cache',
          'upgrade-insecure-requests': '1',
          cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
        },
        referrerPolicy: 'strict-origin-when-cross-origin',
        body: null,
        method: 'GET',
        mode: 'cors',
        credentials: 'include'
      }).then(res => res.arrayBuffer())
    // pdfs.push(pdf)

    // console.log((mappedUnique.indexOf(x))+1+ " of "+mappedUnique.length+ " Labs Done for patient "+visit.PIN+" NO "+c+ " of "+dischargeds.length)

    writeFile(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/IPDexceldis/${br}_${month}_${year}_discharged_drs.xlsx`, Buffer.from(pdf))

    return pdf
    // downloadBlob(pdf, br+"_"+month+"_"+year+ '_discharged_drs.xlsx','application/excel')
  }

  async function PAM() {
    for (const year of years) {
      for (const month of months) {
        console.log(br, month, year, 'PAM')
        const DTF = encodeURIComponent('01 ' + month + ' ' + year)
        const DTT = encodeURIComponent(getlastday(month) + ' ' + month + ' ' + year)

        let controlID = await fetch(loc + '/HISEOD/Areas/EOD/Report/ReportPad654.aspx?processdate=' + DTT + '&IsNet=N', {
          headers: {
            accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
            'accept-language': 'en-US,en;q=0.9,ar;q=0.8',
            'upgrade-insecure-requests': '1',
            cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
          },
          referrer: 'http://130.1.2.27/HISIPBILL/IPBILL/Reports/ViewDischargeReport',
          referrerPolicy: 'strict-origin-when-cross-origin',
          body: null,
          method: 'GET',
          mode: 'cors',
          credentials: 'include'
        }).then(a => a.text()).then(
          res => res.slice(
            res.indexOf('ControlID=') + 'ControlID='.length,
            res.indexOf('&Mode')
          ).substring(0, 32))
        const link = loc + '/HISEOD/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + controlID + '&Mode=true&OpType=Export&FileName=ReportPad654&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML'
        const pdf = await fetch(link
          , {
            headers: {
              accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
              'accept-language': 'en-US,en;q=0.9',
              'cache-control': 'no-cache',
              pragma: 'no-cache',
              'upgrade-insecure-requests': '1',
              cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
            },
            referrerPolicy: 'strict-origin-when-cross-origin',
            body: null,
            method: 'GET',
            mode: 'cors',
            credentials: 'include'
          }).then(res => res.arrayBuffer())
        // pdfs.push(pdf)

        // console.log((mappedUnique.indexOf(x))+1+ " of "+mappedUnique.length+ " Labs Done for patient "+visit.PIN+" NO "+c+ " of "+dischargeds.length)

        writeFile(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/PAD654/${br}_${month}_${year}_ReportPad654.xlsx`, Buffer.from(pdf))

        // downloadBlob(pdf, br+"_"+month+"_"+year+ '_ReportPad654.xlsx','application/excel')

        controlID = null
      }
    }
  }

  async function getBillTable(html, IPIDNo) {
    const ipidss = IPIDNo
    function gettable(table, ClassName, Cate) {
      const tt = []
      table.find('tr').each((i, row) => {
        const rowd = []
        $(row).find('td').each((j, cell) => {
          rowd.push($(cell).text())
        })
        tt.push(rowd)
      })
      // No.', 'Item Code', 'Item Name', 'Order Date', 'Quantity', 'Price', 'Exclusion Price', 'Discount', 'Deductable', 'Total'
      return tt.filter(a => a.length && a.length > 4).map(a => ({
        ipid: ipidss,
        ClassName,
        Category: Cate,
        No: a[0],
        ItemCode: a[1],
        ItemName: a[2],
        OrderDate: a[3],
        Quantity: a[4],
        Price: a[5],
        ExclusionPrice: a[6],
        Discount: a[7],
        Deductible: a[8],
        Total: a[9]

      })).filter(a => !a.Quantity.includes('Sub'))
    }

    // change page to fetch request of IPID

    const page = html
    const $ = cheerio.load(page)
    const x = $('#div-print-detail').children().map((i, el) => {
      const className = $(el).attr('class')
      const Cate = $(el).find('._printheader:first-of-type').text()
      const obj = gettable($(el).find('table'), className, Cate)
      return obj
    }).get()

    // x.IPID=file
    return x
  }

  async function getPage2(html, IPIDt) {
    const ipidss = IPIDt
    function gettable(table, ClassName, Cate, file) {
      const tt = []
      table.find('tr').each((i, row) => {
        const rowd = []
        $(row).find('td').each((j, cell) => {
          rowd.push($(cell).text())
        })
        tt.push(rowd)
      })
      // No.', 'Item Code', 'Item Name', 'Order Date', 'Quantity', 'Price', 'Exclusion Price', 'Discount', 'Deductable', 'Total'
      return tt.filter(a => a.length && a.length > 4).map(a => ({
        ipid: ipidss,
        ClassName,
        Category: Cate,
        No: a[0],
        ItemCode: a[1],
        ItemName: a[2],
        OrderDate: a[3],
        Quantity: a[4],
        Price: a[5],
        Total: a[6]
        // ['Discount']:a[7],
        // ['Deductible']:a[8],
        // ['Total']:a[9],

      })).filter(a => !a.OrderDate.includes('Sub'))
    }

    // change page to fetch request of IPID

    const page = html
    const $ = cheerio.load(page)
    const x = $('#div-print-detail').children().map((i, el) => {
      const className = $(el).attr('class')
      const Cate = $(el).find('._printheader:first-of-type').text()
      const obj = gettable($(el).find('table'), className, Cate)
      return obj
    }).get()
    // console.log(x)
    // x.IPID=file
    return x
  }

  async function getNphiesPl() {
    const list = await fetch(loc + '/HISRCMS/SBSMappingProfileDetails/GetSBSProfileList', {
      headers: {
        accept: 'application/json, text/javascript, */*; q=0.01',
        'accept-language': 'en-US,en;q=0.9',
        'content-type': 'multipart/form-data; boundary=----WebKitFormBoundarycT4HHhlXbx5mAZ2g',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
      },
      referrer: loc + '/HISRCMS/SBSMappingProfileDetails',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: '------WebKitFormBoundarycT4HHhlXbx5mAZ2g--\r\n',
      method: 'POST',
      mode: 'cors',
      credentials: 'include'
    }).then(a => a.json()).then(a => a.ReturnData).then(a => a.map(x => x.Id)).catch(a => console.log(a))
    console.log(list)
    const all = []
    for (const x of list) {
      await fetch(loc + '/HISRCMS/SBSMappingProfileDetails/GetSBSDetails', {
        headers: {
          accept: 'application/json, text/javascript, */*; q=0.01',
          'accept-language': 'en-US,en;q=0.9',
          'content-type': 'multipart/form-data; boundary=----WebKitFormBoundary11wBoNH9RXytP6Wp',
          'x-requested-with': 'XMLHttpRequest',
          cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
        },
        referrer: loc + '/HISRCMS/SBSMappingProfileDetails',
        referrerPolicy: 'strict-origin-when-cross-origin',
        body: '------WebKitFormBoundary11wBoNH9RXytP6Wp\r\nContent-Disposition: form-data; name="IntegrationItemMappingProfileId"\r\n\r\n' + x + '\r\n------WebKitFormBoundary11wBoNH9RXytP6Wp\r\nContent-Disposition: form-data; name="CodeSystem"\r\n\r\n\r\n------WebKitFormBoundary11wBoNH9RXytP6Wp--\r\n',
        method: 'POST',
        mode: 'cors',
        credentials: 'include'
      }).then(a => a.json()).then(a => a.ReturnData).then(a => all.push(a))
      // all.push(c)
    }
    return all
  }

  // console.log(` ${br} Done downloadeing eclaim for ${br} ${months.join(", ")}`)
  async function IPD() {
    const opts = {}
    if (type == 'HCP') {
      for (const year of years) {
        for (const month of months) {
          const disseet = await getdisSheet(year, month)
          const workbook = read(disseet, opts)
          const dislist = utils.sheet_to_json(workbook.Sheets.DischargeReport, opts)
          delete dislist['No.']
          const pins = dislist.map(a => a['Pin No']).filter((value, index, array) => array.indexOf(value) === index)

          console.log(pins)
          let dates = []
          let c = 0
          for (const pin of pins) {
            c++
            if (c % 100 == 0) {
              await page.reload()

              cookie = await page.cookies()
            }
            console.log(br, c, pins.length)
            dates.push(await checkDate(pin.trim()))
          }

          // console.log(dates)
          dates = dates.flat().filter(a => new Date(a.DischargeDateTime).toUTCString().includes(year) && new Date(a.DischargeDateTime).toUTCString().includes(month))
          writeFile(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/IPDCSV/${br}_${month}_${year}_checkDates.json`, JSON.stringify(dates))

          let mergedtable = []
          let cc = 0
          const batchSize = 100
          for (let i = 0; i < dates.length; i += batchSize) {
            const batch = dates.filter(a => a.id != '104678').slice(i, i + batchSize)
            for (const visit of batch) {
              try {
                cc++
                if (cc % 98 == 0) {
                  await page.reload()
                  cookie = await page.cookies()
                }
                console.log(visit.id, dates.length, cc, br, 'IPBILL')
                const visitext = await ipbillF(visit)
                const btable = await getPage2(visitext, visit.id)
                mergedtable.push(btable)
              } catch (error) { }
            }
            const csvv = converter.json2csvAsync(mergedtable.flat(), options)
            writeFile(
              `C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/IPDCSV/${br}_${month}_${year}_IPBILLS_${i / batchSize}.csv`,
              await csvv
            )
            mergedtable = []
          }
        }
      }
    }
  }
  async function IPD2() {
    const opts = {}
    if (type == 'HCP') {
      for (const year of years) {
        for (const month of months) {
          const disseet = await getdisSheet(year, month)
          const workbook = read(disseet, opts)
          const dislist = utils.sheet_to_json(workbook.Sheets.DischargeReport, opts)
          console.log(dislist)
          delete dislist['No.']
          const pins = dislist.filter(a => a.Account.includes('CASH')).map(a => a['Pin No']).filter((value, index, array) => array.indexOf(value) === index).map(a => a.replaceAll(' ', ''))

          console.log(pins)
          const inv = []
          const c = 0
          for (const pin of pins.slice(0, 3)) {
            let billsHeaders = await getbillno(pin)
            billsHeaders = billsHeaders.filter(a => new Date(a.DischargeDateTime).toUTCString().includes(month) && new Date(a.DischargeDateTime).toUTCString().includes(year))
            console.log(billsHeaders)
            const invoice = []
            for (const billsHeader of billsHeaders) {
              const cats = await get_IPBill_Cats(billsHeader.billno)
              console.log(cats)
              for (const cat of cats) {
                let bill = await getIPBillServices(billsHeader.billno, cat.Id)

                bill = bill.map(a => ({
                  IPID: billsHeader, catId: cat.Id, catName: cat.Name, ...a

                })).map((obj) => {
                  const flatObj = flattenObject(obj)
                  return flatObj
                })
                invoice.push(bill)
              }
            }
            inv.push(invoice)
          }
          writeFile('C:/Users/mis1.ryd/adas.json', JSON.stringify(inv))
        }
      }
    }
  }

  async function getPayment(ipid) {
    const html = await fetch(loc + '/HISIPBILL/IPBILL/IPBilling/ViewPayment?ipid=' + ipid + '&_=1687332910850', {
      headers: {
        accept: '*/*',
        'accept-language': 'en-US,en;q=0.9,ar;q=0.8',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
      },
      referrer: loc + '/HISIPBILL/IPBILL/IPBilling/IndexAr',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: null,
      method: 'GET',
      mode: 'cors',
      credentials: 'include'
    }).then(a => a.text())
    // console.log(html)
    const labs = '<html>' + html + '</html'
    const dom = new JSDOM(labs)
    const { document } = dom.window

    const tblResultList = document.getElementsByClassName('_alldep _billprint _billdep')

    const payement = tblResultList.length > 0 ? Array.from(tblResultList).map(a => a.textContent).map(x => x.split('\n').map(a => a.replaceAll(' ', ''))).map(a => ({ Type: a[1], Receipt: a[2], Date: a[3], Mode: a[4], Amount: a[5] })) : null
    const payement2 = document.getElementsByClassName('table table-low-bordered table-hover table-condensed').length > 0 ? Array.from(document.getElementsByClassName('table table-low-bordered table-hover table-condensed')[0].tBodies[0].rows).map(a => a.textContent).map(x => x.split('\n').map(a => a.replaceAll(' ', ''))).map(a => ({ Type: a[1], Receipt: a[2], Date: a[3], Mode: a[4], Amount: a[5] })) : null
    const tt = payement && payement.length ? payement : payement2

    // console.log(tt)
    return tt
  }
  async function IPD3() {
    function getICD(buffer) {
      const workbook = xlsx.readFile(buffer)
      let sheet = workbook.Sheets.IPICDCodes
      sheet = xlsx.utils.sheet_to_json(sheet)
      sheet = sheet.map(a => ({
        PIN: a.__EMPTY,
        Name: a.__EMPTY_3,
        admissionDateTime: moment(a.__EMPTY_10).format('MM/DD/YYYY'),
        dischargeDateTime: moment(a.__EMPTY_11).format('MM/DD/YYYY'),
        ICD: a.__EMPTY_17

      }))
      sheet = sheet.filter(a => a.PIN)
      sheet.forEach(a => {
        a.PIN = +a.PIN.replaceAll(' ', '').split('.')[1]
      })
      return sheet
    }
    function jsonToBill(bill) {
      const Invoice = bill.map(obj => Object.values(obj).filter(a => a != '')).map(a => {
        if (a.length == 5) {
          a.unshift('')
        }
        return a
      }).filter(a => a.length == 6 || a.length == 1)
        .filter(a => !a.join('|||').includes('AMOUNT') &&
          !a.join('|||').includes('DISCOUNTABLE') &&
          !a.join('|||').includes('IN WORDS') &&
          !a.join('|||').includes('LESS DEDUCTABLE') &&
          !a.join('|||').includes('LESS DISCOUNT') &&
          !a.join('|||').includes('----') && !a.join('|||').includes('Saudi German') && !a.join('|||').includes('DISCHARGE TIME') && !a.join('|||').includes('IPCR'))
      return Invoice
    }
    const addCategoryToItems = (array) => {
      let currentCategory = ''

      for (let i = 0; i < array.length; i++) {
        const item = array[i]

        if (item.length === 1) {
          // Category item
          currentCategory = item[0]
        } else {
          // Add category to the item
          item.unshift(currentCategory)
        }
      }

      return array.filter(a => a.length > 1)
    }
    function IPinvoicetoJson(loc) {
      const workbook = xlsx.readFile(loc).Sheets.IPInvoiceBreakUpNPD
      const bill = xlsx.utils.sheet_to_json(workbook)
      const invoiceNo = bill.map(obj => Object.values(obj).filter(a => a != '')).map(x => x.join('')).find(x => x.includes('IPCR'))

      const Invoice = bill.map(obj => Object.values(obj).filter(a => a != '')).map(a => {
        if (a.length == 5) {
          a.unshift('')
        }
        return a
      }).filter(a => a.length == 6 || a.length == 1)
        .filter(a => !a.join('|||').includes('AMOUNT') &&
          !a.join('|||').includes('DISCOUNTABLE') &&
          !a.join('|||').includes('IN WORDS') &&
          !a.join('|||').includes('LESS DEDUCTABLE') &&
          !a.join('|||').includes('LESS DISCOUNT') &&
          !a.join('|||').includes('----') && !a.join('|||').includes('Saudi German') && !a.join('|||').includes('DISCHARGE TIME') && !a.join('|||').includes('IPCR'))

      const orignal = Invoice.map((a, index) => [index, a].flat())
      const orinvoiceonly = orignal.filter(a => a.length == 7).map(a => ({
        index: a[0],
        Code: a[1],
        Desc: a[2],
        Date: a[3],
        Qty: a[4],
        Unit: a[5],
        Gross: a[6]

      }))
      let cats = orignal.filter(a => a.length == 2 && typeof (a[1]) !== 'number')

      cats = cats.map(a => ({ start: a[0], end: cats.find(x => x[0] > a[0]) ? cats.find(x => x[0] > a[0])[0] : orignal.length, Category: a[1] }))
      cats.map(a => ({ cat: a.Category, inv: orinvoiceonly.slice(a.start - 1, a.end) })).map(a => a.inv.map(x => ({ cat: a.cat, ...x }))).flat()
      // Invoice=Invoice.map(a=>a.join(",")).join("\n")
      // let csvv = converter.json2csv(Invoice.flat(), options);
      // const resultArray = addCategoryToItems(Invoice).filter(a=>a.length>1);
      // console.log(resultArray)
      // Invoice=resultArray.map(a=>a.join(",")).join("\n")

      const billt = addCategoryToItems(jsonToBill(bill))
      let disc = bill.map(obj => Object.values(obj).filter(a => a != '')).filter(a => a.length == 3 && a.join('').includes('%'))
      let billob = billt.map(a => ({ Category: a[0], Code: a[1], Description: a[2], Date: a[3], Qty: a[4], Unit: a[5], Gross: a[6] }))
      disc = disc.map(a => ({ Category: a[1], Amount: a[2], BillTotal: billob.filter(x => x.Category == a[1]).map(a => a.Gross).reduce((a, b) => a + b) })).map(a => ({ Perc: a.Amount / a.BillTotal, ...a }))
      billob = billob.map(a => ({ Discount: disc.find(x => x.Category == a.Category) ? disc.find(x => x.Category == a.Category).Perc * a.Gross : 0, ...a }))
      billob.forEach(a => a.Invoicenumber = invoiceNo)
      // writeFile('C:/Users/mis1.ryd/asd.json',JSON.stringify(billob))
      // writeFile('C:/Users/mis1.ryd/asd.csv',convertJSONtoCSV(billob))
      // console.log(billob.slice(0,10),disc)
      return billob
    }

    const opts = {}
    if (type == 'HCP') {
      const cash = await cashlist()
      // console.log(cash)
      for (const year of years) {
        for (const month of months) {
          const jsons = []
          const invoices = []
          const errs = []
          const pays = []
          const disseet = await getdisSheet(year, month)
          let ICD = await censusIPD(year, month)
          ICD = getICD(ICD)
          // console.log(ICD)
          const workbook = read(disseet, opts)
          const dislist = utils.sheet_to_json(workbook.Sheets.DischargeReport, opts)
          // console.log( dislist[0] )
          delete dislist['No.']
          const pins = dislist.filter(a => cash.list.includes(a.Account))
            .map(a => a['Pin No']).filter((value, index, array) => array.indexOf(value) === index).map(a => a.replaceAll(' ', ''))

          // console.log(pins)

          let c = 0
          for (const pin of pins) {
           // console.log(pin)
            try {
              if (c % 100 == 0) {
                await page.reload()

                cookie = await page.cookies()
              }

              c++
              const ptdt = await ptDetails(pin)
              let billsHeaders = await getbillno(pin)
              billsHeaders=getUniqueByBillNo(billsHeaders)
              jsons.push({ PIN: pin, visits: billsHeaders })
              billsHeaders = billsHeaders.filter(a => a.CategoryId == cash.comps && moment(a.DischargeDateTime).format('MMM-YYYY').includes(month) && moment(a.DischargeDateTime).format('MMM-YYYY').includes(year))
              console.log(Math.round(c * 100 / pins.length), c, pins.length, br, 'CashIPD', month, year)
              //console.log(billsHeaders)
              for (const billsHeader of billsHeaders) {
                // try {

                const pay = await getPayment(billsHeader.id)
                if (pay && pay.length > 0) {
                  pay.forEach(a => {
                    a.IPID = billsHeader.id,
                      a.Branch = br
                  })
                  pays.push(pay)
                }
                // console.log(pay)
                const pdf = await getbillexcel(billsHeader.billno)
                let invoicejson = IPinvoicetoJson(pdf)
              //  console.log(invoicejson.length, 'invoicejson1')
                const dislistkey = dislist.find(
                  a =>
                    a['Pin No'].replaceAll(' ', '') == pin &&
                    ExcelDateToJSDate(a['Admit Date Time']).toISOString().split('T')[0] == new Date(billsHeader.AdmitDateTime).toISOString().split('T')[0] &&
                    ExcelDateToJSDate(a['Discharge Date Time']).toISOString().split('T')[0] == new Date(billsHeader.DischargeDateTime).toISOString().split('T')[0]

                )

                invoicejson.forEach(a => { a.header = billsHeader })
               // console.log(invoicejson.length, 'invoicejson2')
                invoicejson = invoicejson.map(({ header, ...rest }) => {
                  return { ...header, ...rest }
                })
               // console.log(invoicejson.length, 'invoicejson3')
                invoicejson.forEach(a => { a.header = dislistkey })
                invoicejson = invoicejson.map(({ header, ...rest }) => {
                  return { ...header, ...rest }
                })
               // console.log(invoicejson.length, 'invoicejson4')
                let ICDs = ICD.find(a => a.PIN == +pin && moment(billsHeader.AdmitDateTime).format('MM/DD/YYYY') == a.admissionDateTime && moment(billsHeader.DischargeDateTime).format('MM/DD/YYYY') == a.dischargeDateTime)
                ICDs = ICDs ? ICDs.ICD : null

                invoicejson.forEach(a => { a.ICDCODE = ICDs })
               // console.log(invoicejson.length, 'invoicejson5')

                const eclaimformat = invoicejson.map(a => ({

                  PIN: a['Pin No'],
                  PTNAME: a['Patient Name'],
                  MEDID: ptdt.PTDetails.ResidenceId,
                  COMPANY: a.Account,
                  POLICYNO: '',
                  INVOICENO: a.Invoicenumber,
                  ICDCODE: a.ICDCODE,
                  ICDDESC: ' ',
                  BILLDATETIME: a.Date,
                  SERVICECODE: a.Code,
                  SERVICENAME: a.Description,
                  QTY: a.Qty,
                  GROSS: a.Gross,
                  DISCOUNT: a.Discount,
                  DEDUCTIBLE: '0',
                  NET: +a.Gross - (+a.Discount),
                  IQAMA: ptdt.PTDetails.ResidenceId,
                  VAT: '0',
                  ADMITDATE: a.AdmitDateTime,
                  DISCHARGEDATE: a.DischargeDateTime

                }))
                console.log(eclaimformat.length, 'eclaimformat')
                // console.log(invoicejson[0])
                invoices.push(eclaimformat)
                //console.log(invoices.length, 'invoices')

                // let workbook2 = read(disseet, opts)
                // let dislist1 = utils.sheet_to_json(workbook2.Sheets.DischargeReport, opts)

                // writeFile("C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Cashwallahy/"+br+"_"+billsHeader.billno+".xlsx", Buffer.from(pdf))
              }
            } catch (error) {
              console.log(error)
              errs.push(pin)
            }
          }

          jsons.length ? fs.writeFileSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Cashwallahy/' + br + '_' + month + '_' + year + '_IPBILL.json', JSON.stringify(jsons)) : null
          console.log(invoices.length)
          let csv = converter.json2csvAsync(invoices.flat(), options)
          invoices.length ? fs.writeFileSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Cashwallahy/' + br + '_' + month + '-' + year + '_' + cash.comps + '_1_IP.csv', await csv) : null
          writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Cashwallahy/' + br + '_' + month + '_' + year + '_errs.json', JSON.stringify(errs))
          pays.length ? fs.writeFileSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Cashwallahy/' + br + '_' + month + '-' + year + '_' + cash.comps + '_1_Payement.csv', convertJSONtoCSV(pays.flat())) : null
        }
      }
    }
  }
  async function IPD4() {
    const opts = {}
    if (type == 'HCP') {
      for (const year of years) {
        for (const month of months) {
          const jsons = []
          const disseet = await getdisSheet(year, month)
          const workbook = read(disseet, opts)
          const dislist = utils.sheet_to_json(workbook.Sheets.DischargeReport, opts)
          // console.log(dislist)
          delete dislist['No.']
          const pins = dislist.map(a => a['Pin No']).filter((value, index, array) => array.indexOf(value) === index).map(a => a.replaceAll(' ', ''))

          // console.log(pins)

          let c = 0
          for (const pin of pins) {
            if (c % 100 == 0) {
              await page.reload()

              cookie = await page.cookies()
            }

            c++
            const billsHeaders = await getbillno(pin)
            jsons.push({ PIN: pin, visits: billsHeaders })
            console.log(jsons.length, br, 'IPBILLS json')
          }

          writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Cashwallahy/' + br + '_' + month + '_' + year + '_IPBILL.json', JSON.stringify(jsons))
        }
      }
    }
  }

  async function BillingEff(year, month) {
    console.log(br, month, year, 'effec')

    await fetch(loc + '/HISMCRS/ManagementReports/ARReports/BillingEfficiencyReport', {
      headers: {
        accept: '*/*',
        'accept-language': 'en-US,en;q=0.9',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
      },
      referrer: loc + '/HISMCRS/ManagementReports/ARReports/BillingEfficiencyReport',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: 'CategoryId=0&StartDate=01-' + month + '-' + year + '&EndDate=' + getlastday(month) + '-' + month + '-' + year + '&X-Requested-With=XMLHttpRequest',
      method: 'POST',
      mode: 'cors',
      credentials: 'include'
    })
    const controlID = await fetch(loc + '/HISMCRS/ReportViewerWebForm.aspx', {
      headers: {
        accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'accept-language': 'en-US,en;q=0.9',
        'upgrade-insecure-requests': '1',
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
      },
      referrer: loc + '/HISMCRS/ManagementReports/ARReports/BillingEfficiencyReport',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: null,
      method: 'GET',
      mode: 'cors',
      credentials: 'include'
    }).then(a => a.text()).then(
      res => res.slice(
        res.indexOf('ControlID=') + 'ControlID='.length,
        res.indexOf('&Mode')
      ).substring(0, 32))
    const link = loc + '/HISMCRS/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + controlID + '&Mode=true&OpType=Export&FileName=BillingEfficiency&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML'

    // http://130.8.2.18/HISMCRS/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=f9ae161c5d9649358f7db35c396ee0bf&Mode=true&OpType=Export&FileName=BillingEfficiency&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML
    const pdf = await fetch(link
      , {
        headers: {
          accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
          'accept-language': 'en-US,en;q=0.9',
          'cache-control': 'no-cache',
          pragma: 'no-cache',
          'upgrade-insecure-requests': '1',
          cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
        },
        referrerPolicy: 'strict-origin-when-cross-origin',
        body: null,
        method: 'GET',
        mode: 'cors',
        credentials: 'include'
      }).then(res => res.arrayBuffer())
    // pdfs.push(pdf)

    // console.log((mappedUnique.indexOf(x))+1+ " of "+mappedUnique.length+ " Labs Done for patient "+visit.PIN+" NO "+c+ " of "+dischargeds.length)

    writeFile(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Effec/${br}_${month}_${year}_BE.xlsx`, Buffer.from(pdf))

    // return pdf
    // downloadBlob(pdf, br+"_"+month+"_"+year+ '_discharged_drs.xlsx','application/excel')
  }

  async function censusOPD(tt, ttz) {
    const downloaded = fs.readdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/DailyCensus/').filter(a => a.includes(br)&&!a.includes("dailyCensusIP")).map(a => a.replace('_dailyCensus', '').replace(br + '_', '').replace('.xlsx', '').replaceAll('_', '/'))
    console.log(br, downloaded.length)
    function generateDateArray(startDate, endDate) {
      const dateArray = []
      const currentDate = new Date(startDate)

      while (currentDate <= endDate) {
        const month = (currentDate.getMonth() + 1).toString().padStart(2, '0')
        const day = currentDate.getDate().toString().padStart(2, '0')
        const year = currentDate.getFullYear().toString()
        const formattedDate = month + '/' + day + '/' + year

        dateArray.push(formattedDate)
        currentDate.setDate(currentDate.getDate() + 1)
      }

      return dateArray
    }
    const startDate = new Date(tt)
    const endDate = new Date(ttz)

    const dates = generateDateArray(startDate, endDate).reverse().filter(a => !downloaded.includes(a))
    console.log(br, dates.length, 'rem')
    for (const date of dates) {
      console.log(date, br)
      const res = await fetch(loc + '/HISMRD2/Reports/MRDReportViewer.aspx?rpt=25&FromDate=' + date + '&ToDate=' + date + '&NationalityID=0&DepartmentID=0&DoctorID=0&SexID=0&ICD=0&CompanyID=0', {
        headers: {
          accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
          'accept-language': 'en-US,en;q=0.9',
          'upgrade-insecure-requests': '1',
          cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
        },
        referrer: 'http://130.1.2.27/HISMRD2/MRD/Reports/OP_Patients_ICD_Code/',
        referrerPolicy: 'strict-origin-when-cross-origin',
        body: null,
        method: 'GET',
        mode: 'cors',
        credentials: 'include'
      }).then(a => a.text())
      const ctrl = res.slice(
        res.indexOf('ControlID=') + 'ControlID='.length,
        res.indexOf('&Mode')
      ).substring(0, 32)
      const link = loc + '/HISMRD2/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + ctrl + '&Mode=true&OpType=Export&FileName=OP_Patients_ICD_Code&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML'
      const pdf = await fetch(link, {
        headers: {
          accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
          'accept-language': 'en-US,en;q=0.9',
          'cache-control': 'no-cache',
          pragma: 'no-cache',
          'upgrade-insecure-requests': '1',
          cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')

        },
        referrerPolicy: 'strict-origin-when-cross-origin',
        body: null,
        method: 'GET',
        mode: 'cors',
        credentials: 'include'
      }).then(a => a.arrayBuffer())

      // fs.existsSync(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/DailyCensus/`)?null:fs.mkdirSync(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/DailyCensus/`)
      writeFile(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/DailyCensus/${br}_${date.replaceAll('/', '_')}_dailyCensus.xlsx`, Buffer.from(pdf))
    }
  }

  async function Registeration(fromx, tox) {
    const from = fromx
    const to = tox
    const period = getMonthStartAndEndDates2(from, to).reverse()

    for (const x of period) {
      console.log(period.to, br, 'Registeration')
      const downloaded = []
      // fs.readdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Register/').filter(a=>a.includes(br)).map(a=>a.replace("_Register","").replace(br+"_","").replace(".xlsx","").replaceAll("_","/"))
      console.log(br, downloaded.length)
      function generateDateArray(startDate, endDate) {
        const dateArray = []
        const currentDate = new Date(startDate)

        while (currentDate <= endDate) {
          const month = (currentDate.getMonth() + 1).toString().padStart(2, '0')
          const day = currentDate.getDate().toString().padStart(2, '0')
          const year = currentDate.getFullYear().toString()
          const formattedDate = month + '/' + day + '/' + year

          dateArray.push(formattedDate)
          currentDate.setDate(currentDate.getDate() + 1)
        }

        return dateArray
      }
      const fromm = '2023-01-01'
      const too = '2023-03-01'
      const startDate = new Date(x.start)
      const endDate = new Date(x.end)

      const dates = generateDateArray(startDate, endDate).reverse().filter(a => !downloaded.includes(a))
      console.log(br, dates.length, 'rem')
      let arr = []
      for (const date of dates) {
        try {
          console.log(date, br)
          const res = await fetch(loc + '/HISMRD2/Reports/MRDReportViewer.aspx?rpt=1&FromDate=' + date + '&ToDate=' + date, {
            headers: {
              accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
              'accept-language': 'en-US,en;q=0.9',
              'upgrade-insecure-requests': '1',
              cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
            },
            referrer: 'http://130.1.2.27/HISMRD2/MRD/Reports/OP_Patients_ICD_Code/',
            referrerPolicy: 'strict-origin-when-cross-origin',
            body: null,
            method: 'GET',
            mode: 'cors',
            credentials: 'include'
          }).then(a => a.text())
          const ctrl = res.slice(
            res.indexOf('ControlID=') + 'ControlID='.length,
            res.indexOf('&Mode')
          ).substring(0, 32)
          // http://130.1.2.153/HISMRD2/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=85ad41348567477596ae9c942b5a1aa0&Mode=true&OpType=Export&FileName=Registrations&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML
          const link = loc + '/HISMRD2/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + ctrl + '&Mode=true&OpType=Export&FileName=Registrations&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML'
          const pdf = await fetch(link, {
            headers: {
              accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
              'accept-language': 'en-US,en;q=0.9',
              'cache-control': 'no-cache',
              pragma: 'no-cache',
              'upgrade-insecure-requests': '1',
              cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')

            },
            referrerPolicy: 'strict-origin-when-cross-origin',
            body: null,
            method: 'GET',
            mode: 'cors',
            credentials: 'include'
          }).then(a => a.arrayBuffer())
          const opts = {}
          const file = xlsx.readFile(pdf, opts)
          let json = xlsx.utils.sheet_to_json(file.Sheets.Registrations)
          json = json.map(a => ({ Branch: br, date, Pin: a.__EMPTY_2, Name: a.__EMPTY_7, Age: a.__EMPTY_11, Gender: a.__EMPTY_12, City: a.__EMPTY_13 }))
          arr.push(json)

          // fs.existsSync(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/DailyCensus/`)?null:fs.mkdirSync(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/DailyCensus/`)
          // writeFile(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Register/${br}_${date.replaceAll('/',"_")}_Register.xlsx`, Buffer.from(pdf))
        } catch (error) {

        }
      }
      arr = arr.flat()

      writeFile(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Register/${br}_${x.start}_${x.end}_Register.csv`, convertJSONtoCSV(arr))
    }
  }
  async function census3() {
    // Example usage
    const startDate = '2023-06-01' // Start of the period
    const endDate = '2023-07-01' // End of the period

    const monthDates = getMonthStartAndEndDates(startDate, endDate).reverse()
    console.log(monthDates)
    monthDates
    for (const x of monthDates) {
      console.log(x.end, br)
      /// HISMRD2/Reports/MRDReportViewer.aspx?rpt=30&FromDate="+x.start+"%20&ToDate=%20"+x.end+"&Mode=2&NationalityID=0&DepartmentId=0&DoctorID=0&SexID=0&ICD=0&CompanyID=0&CodeType=3
      const res = await fetch(loc + '/HISMRD2/Reports/MRDReportViewer.aspx?rpt=30&FromDate=' + x.start + '%20&ToDate=%20' + x.end + '&Mode=2&NationalityID=0&DepartmentId=0&DoctorID=0&SexID=0&ICD=0&CompanyID=0&CodeType=3', {
        headers: {
          accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
          'accept-language': 'en-US,en;q=0.9',
          'upgrade-insecure-requests': '1',
          cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
        },
        referrer: 'http://130.1.2.27/HISMRD2/MRD/Reports/OP_Patients_ICD_Code/',
        referrerPolicy: 'strict-origin-when-cross-origin',
        body: null,
        method: 'GET',
        mode: 'cors',
        credentials: 'include'
      }).then(a => a.text())
      const ctrl = res.slice(
        res.indexOf('ControlID=') + 'ControlID='.length,
        res.indexOf('&Mode')
      ).substring(0, 32)
      const link = loc + '/HISMRD2/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + ctrl + '&Mode=true&OpType=Export&FileName=OP_Patients_ICD_Code&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML'
      const pdf = await fetch(link, {
        headers: {
          accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
          'accept-language': 'en-US,en;q=0.9',
          'cache-control': 'no-cache',
          pragma: 'no-cache',
          'upgrade-insecure-requests': '1',
          cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')

        },
        referrerPolicy: 'strict-origin-when-cross-origin',
        body: null,
        method: 'GET',
        mode: 'cors',
        credentials: 'include'
      }).then(a => a.arrayBuffer())

      // fs.existsSync(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/DailyCensus/`)?null:fs.mkdirSync(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/DailyCensus/`)
      writeFile(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/DailyCensus/${br}_${x.end.replaceAll('/', '_')}_dailyCensus.xlsx`, Buffer.from(pdf))
    }
  }

  async function Census(from, to) {
    const days = datesInPeriod(from, to)
    console.log(days)

    for (const day of days) {
      {
        const req = await fetch(loc + '/HISMCRS/ManagementReports/FinanceReports/PatientCensusReport', {
          headers: {
            accept: '*/*',
            'accept-language': 'en-US,en;q=0.9,ar;q=0.8',
            'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'x-requested-with': 'XMLHttpRequest',
            cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
          },
          referrer: loc + '/HISMCRS/ManagementReports/FinanceReports/PatientCensusReport',
          referrerPolicy: 'strict-origin-when-cross-origin',
          body: `ReportTypeId=1&FromDate=${day}&ToDate=${day}&X-Requested-With=XMLHttpRequest`,
          method: 'POST',
          mode: 'cors',
          credentials: 'include'
        }).then(a => a.text()).then(a => !a.includes('NO RECORDS FOUND'))
        if (req)
        // console.log(day)
        {
          let E6 = null
          let ctrlid = null
          let buff = null
          let issheet = null
          while (!ctrlid || !buff || !issheet || !E6) {
            console.log(`${br}-${day}`)

            const response = await fetch(loc + '/HISMCRS/ReportViewerWebForm.aspx', {
              headers: {
                accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
                'accept-language': 'en-US,en;q=0.9,ar;q=0.8',
                'cache-control': 'max-age=0',
                'upgrade-insecure-requests': '1',
                cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
              },
              referrerPolicy: 'strict-origin-when-cross-origin',
              body: null,
              method: 'GET',
              mode: 'cors',
              credentials: 'include'
            })

            const text = await response.text()
            const ControlID = text.slice(
              text.indexOf('ControlID=') + 'ControlID='.length,
              text.indexOf('&Mode')
            ).substring(0, 32)

            if (text.includes('PatientCensusReport') && !ControlID.includes('html')) {
              // console.log(ControlID)

              ctrlid = ControlID
              const xx = await fetch(loc + '/HISMCRS/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + ControlID + '&Mode=true&OpType=Export&FileName=PatientCensusReport_ByCategoryCompany&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML'
                , {
                  headers: {
                    accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                    'accept-language': 'en-US,en;q=0.9',
                    'cache-control': 'no-cache',
                    pragma: 'no-cache',
                    'upgrade-insecure-requests': '1',
                    cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
                  },
                  referrerPolicy: 'strict-origin-when-cross-origin',
                  body: null,
                  method: 'GET',
                  mode: 'cors',
                  credentials: 'include'
                })
              buff = xx.ok
              // console.log(budd)
              const orrep = await xx.arrayBuffer()
              const m = day

              E6 = read(orrep).Sheets.PatientCensusReport_ByCategoryC.E6.v
              issheet = E6.includes(m)
              console.log(br, m, E6, issheet)

              writeFile(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Eclaim/Census/${br}_${moment(day, 'DD-MMM-YYYY').format('DD_MMM_YYYY')}_Census.xlsx`, Buffer.from(orrep))
            }
          }
          ctrlid = null
          buff = null
          issheet = false
        }
      }

      // throw new Error(`Fetch failed after ${maxAttempts} attempts`);
    }
  }

  async function DoctorSchedule(from, to) {
    //01+Jan+2022



    await fetch(loc + "/HISRECEPTION/Reception/Reports/DoctorShiftSchedule", {
      "headers": {
        "accept": "*/*",
        "accept-language": "en-US,en;q=0.9,ar;q=0.8",
        "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
        "x-requested-with": "XMLHttpRequest",
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
      },
      "referrer": "http://130.1.2.27/HISRECEPTION/Reception/Reports/DoctorShiftSchedule",
      "referrerPolicy": "strict-origin-when-cross-origin",
      "body": "StartDate=" + from + "&EndDate=" + to + "&DoctorId=&GroupByDoctor=false&X-Requested-With=XMLHttpRequest",
      "method": "POST",
      "mode": "cors",
      "credentials": "include"
    });


    console.log(`${br}-${from} Dr schedule`)

    const response = await fetch(loc + "/HISRECEPTION/ReportViewerWebForm.aspx", {
      "headers": {
        "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "accept-language": "en-US,en;q=0.9,ar;q=0.8",
        "upgrade-insecure-requests": "1",
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
      },
      "referrer": loc + "/HISRECEPTION/Reception/Reports/DoctorShiftSchedule",
      "referrerPolicy": "strict-origin-when-cross-origin",
      "body": null,
      "method": "GET",
      "mode": "cors",
      "credentials": "include"
    })

    const text = await response.text()
    const ControlID = text.slice(
      text.indexOf('ControlID=') + 'ControlID='.length,
      text.indexOf('&Mode')
    ).substring(0, 32)

    const xx = await fetch(loc + "/HISRECEPTION/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=" + ControlID + "&Mode=true&OpType=Export&FileName=DoctorShiftDetailsByDoctor&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML"
      , {
        headers: {
          accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
          'accept-language': 'en-US,en;q=0.9',
          'cache-control': 'no-cache',
          pragma: 'no-cache',
          'upgrade-insecure-requests': '1',
          cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
        },
        referrerPolicy: 'strict-origin-when-cross-origin',
        body: null,
        method: 'GET',
        mode: 'cors',
        credentials: 'include'
      })

    const orrep = await xx.arrayBuffer()



    writeFile(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/DoctorSchedule/${br}_${from}_${to}_DoctorSchedule.xlsx`, Buffer.from(orrep))
  }



  async function CensusDr() {
    for (const year of years) {
      for (const month of months) {
        const startofmonth = new Date('01 ' + month + ' 2022')
        const lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 0)
        const dst = lastDayOfMonth.toISOString().split('T')[0]
        const lstdd = +dst.substr(dst.length - 2) + 1
        for (let day = 1; new Date(year & '-' & month & '-' & day) <= new Date() && day <= lstdd; day++) {
          if (day == lstdd || true) {
            const req = await fetch(loc + '/HISMCRS/ManagementReports/FinanceReports/PatientCensusReport', {
              headers: {
                accept: '*/*',
                'accept-language': 'en-US,en;q=0.9,ar;q=0.8',
                'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
                'x-requested-with': 'XMLHttpRequest',
                cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
              },
              referrer: loc + '/HISMCRS/ManagementReports/FinanceReports/PatientCensusReport',
              referrerPolicy: 'strict-origin-when-cross-origin',
              body: `ReportTypeId=3&FromDate=${day}-${month}-${year}&ToDate=${day}-${month}-${year}&X-Requested-With=XMLHttpRequest`,
              method: 'POST',
              mode: 'cors',
              credentials: 'include'
            }).then(a => a.text()).then(a => !a.includes('NO RECORDS FOUND'))
            if (req)
            // console.log(day)
            {
              let ctrlid = null
              let buff = null
              let issheet = false
              while (!ctrlid || !buff || !issheet) {
                // console.log(`${br}-${day}-${month}-${year}`)

                const response = await fetch(loc + '/HISMCRS/ReportViewerWebForm.aspx', {
                  headers: {
                    accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
                    'accept-language': 'en-US,en;q=0.9,ar;q=0.8',
                    'cache-control': 'max-age=0',
                    'upgrade-insecure-requests': '1',
                    cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
                  },
                  referrerPolicy: 'strict-origin-when-cross-origin',
                  body: null,
                  method: 'GET',
                  mode: 'cors',
                  credentials: 'include'
                })

                const text = await response.text()
                const ControlID = text.slice(
                  text.indexOf('ControlID=') + 'ControlID='.length,
                  text.indexOf('&Mode')
                ).substring(0, 32)

                if (text.includes('PatientCensusReport') && !ControlID.includes('html')) {
                  // console.log(ControlID)

                  ctrlid = ControlID
                  const xx = await fetch(loc + '/HISMCRS/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + ControlID + '&Mode=true&OpType=Export&FileName=PatientCensusReport_ByCategoryCompany&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML'
                    , {
                      headers: {
                        accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                        'accept-language': 'en-US,en;q=0.9',
                        'cache-control': 'no-cache',
                        pragma: 'no-cache',
                        'upgrade-insecure-requests': '1',
                        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
                      },
                      referrerPolicy: 'strict-origin-when-cross-origin',
                      body: null,
                      method: 'GET',
                      mode: 'cors',
                      credentials: 'include'
                    })
                  buff = xx.ok
                  // console.log(budd)
                  const orrep = await xx.arrayBuffer()
                  const m = moment(new Date(`${month}-${day}-${year}`)).format('DD-MMM-yyyy')
                  // console.log(m)
                  const E6 = read(orrep).Sheets.PatientCensusReport_ByDoctor.F6.v
                  issheet = E6.includes(m)
                  console.log(br, m, E6, issheet)

                  writeFile(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Eclaim/CensusDr/${br}_${day}_${month}_${year}_CensusDr.xlsx`, Buffer.from(orrep))
                }
              }
              ctrlid = null
              buff = null
              issheet = false
            }

            // throw new Error(`Fetch failed after ${maxAttempts} attempts`);
          }
        }
      }
    }
  }

  async function getDis(year, month) {
    console.log(month, year, 'Dischargejson')
    const xar = [{ m: 'Jan', num: '01' }, { m: 'Feb', num: '02' }, { m: 'Mar', num: '03' },
    { m: 'Apr', num: '04' }, { m: 'May', num: '05' }, { m: 'Jun', num: '06' },
    { m: 'Jul', num: '07' }, { m: 'Aug', num: '08' }, { m: 'Sep', num: '09' },
    { m: 'Oct', num: '10' }, { m: 'Nov', num: '11' }, { m: 'Dec', num: '12' }]
    month == xar.find(a => a.m == month).num
    const startofmonth = new Date('01 ' + month + ' ' + year)
    const lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 0)
    const dst = lastDayOfMonth.toISOString().split('T')[0]
    const lstdd = +dst.substr(dst.length - 2) + 1
    console.log(month)
    const discharged = await fetch(loc + '/HISARADMIN/ARBillFinalization/getDischargeBills', {
      headers: {
        accept: 'application/json, text/javascript, */*; q=0.01',
        'accept-language': 'en-US,en;q=0.9',
        'cache-control': 'no-cache',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        pragma: 'no-cache',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
      },
      referrer: loc + '/HISARADMIN/ARBillFinalization',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: 'categoryId=0&fDate=01-' + month + '-' + year + '&tDate=' + lstdd + '-' + month + '-' + year,
      method: 'POST',
      mode: 'cors',
      credentials: 'include'
    }).then((response) => response.json())
    writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/VBHC/' + br + '_' + month + '_' + year + '_discharged.json', JSON.stringify(discharged))

    // downloadBlob(JSON.stringify(discharged), br + "_" + month + "_" + year + '_discharged.json', 'text/html')
    const dischargeds = discharged.Res// .sort(compareAge)
    return dischargeds
  }

  async function progressnotes(dischargeds, year, month) {
    // let dischargeds= await getDis()
    const prs = []
    let c = 0
    for (const x of dischargeds) {
      c++
      if (c % 50 == 0 && c > 49) {
        await page.reload()

        cookie = await page.cookies()
      }
      console.log('Downloading Progress Notes for patinet ' + dischargeds.indexOf(x) + ' of ' + dischargeds.length, br)
      const prnotes = await fetch(loc + '/' + dm + '/DM/Generic/DM_VW_PROGGNOT?ipid=' + x.IPID + '&vid=0&reg=' + (+x.PIN.split('.').pop()) + '&_=1669183618988', {
        headers: {
          accept: '*/*',
          'accept-language': 'en-US,en;q=0.9',
          'cache-control': 'no-cache',
          pragma: 'no-cache',
          'x-requested-with': 'XMLHttpRequest',
          cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
        },
        referrer: loc + '/HISWARDS/Wards/WorkOrder/WorkOrder',
        referrerPolicy: 'strict-origin-when-cross-origin',
        body: null,
        method: 'GET',
        mode: 'cors',
        credentials: 'include'
      }).then((response) => response.text())

      prs.push({ mrn: x.PIN.split('.').pop(), IPID: x.IPID, html: prnotes })
    }
    writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/VBHC/' + br + '_' + month + '_' + year + '_pr_' + '.html', JSON.stringify(prs))

    //  downloadBlob(JSON.stringify(prs), br + "_" + month + "_" + year + "_pr_" + '.html', 'text/html')
  }

  async function dateWiseSurgey(year, month) {
    // var srst="01/01/2023&ToDate=01/25/2023"
    const startofmonth = new Date('01 ' + month + ' ' + year)
    const lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 0)
    const dst = lastDayOfMonth.toISOString().split('T')[0]
    const lstdd = +dst.substr(dst.length - 2) + 1
    const srst = year + '-' + month + '-01' + '&ToDate=' + year + '-' + month + '-' + lstdd
    console.log('Downloading Date Wise surgey for ' + br + '_' + month)

    const res = await fetch(loc + '/HISMRD2/Reports/MRDReportViewer.aspx?rpt=55&CompanyID=0&FromDate=' + srst + '&OrNotes=1&DepartmentID=0&DoctorID=0', {
      headers: {
        accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'accept-language': 'en-US,en;q=0.9',
        'upgrade-insecure-requests': '1',
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
      },
      referrer: loc + '/HISMRD2/MRD/Reports/Date_wise_Surgeries/',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: null,
      method: 'GET',
      mode: 'cors',
      credentials: 'include'
    }).then((r) => r.text())

    const ControlID = res.slice(
      res.indexOf('ControlID=') + 'ControlID='.length,
      res.indexOf('&Mode')
    ).substring(0, 32)
    // downloadBlob(ControlID,'res.html','text/html')
    // window.alert( ControlID.substring(0,32))
    const orrep = await fetch(loc + '/HISMRD2/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + ControlID + '&Mode=true&OpType=Export&FileName=Get_Date_Wise_Surgeries&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML'
      , {
        headers: {
          accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
          'accept-language': 'en-US,en;q=0.9',
          'cache-control': 'no-cache',
          pragma: 'no-cache',
          'upgrade-insecure-requests': '1',
          cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
        },
        referrerPolicy: 'strict-origin-when-cross-origin',
        body: null,
        method: 'GET',
        mode: 'cors',
        credentials: 'include'
      }).then(res => res.arrayBuffer())

    writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/VBHC/ORsheets/' + br + '_' + month + '_' + year + '_Date_Wise_surgery_IP.xlsx', Buffer.from(orrep))

    // downloadBlob(orrep, br + "_" + month + "_Date_Wise_surgery_IP.xlsx", 'application/excel')
  }

  async function ORIPRep(year, month) {
    // var srst="01/01/2023&ToDate=01/25/2023"
    const startofmonth = new Date('01 ' + month + ' ' + year)
    const lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 0)
    const dst = lastDayOfMonth.toISOString().split('T')[0]
    const lstdd = +dst.substr(dst.length - 2) + 1
    const srst = year + '-' + month + '-01' + '&ToDate=' + year + '-' + month + '-' + lstdd
    console.log('Downloading OR IP orders for ' + br + '_' + month)

    const res = await fetch(loc + '/HISMRD2/Reports/MRDReportViewer.aspx?rpt=50&FromDate=' + srst + '&Asst1ID=0&Asst2ID=0&DoctorID=0&DepartmentID=0&SugeryID=0&NationalityID=0', {
      headers: {
        accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'accept-language': 'en-US,en;q=0.9',
        'upgrade-insecure-requests': '1',
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
      },
      referrer: loc + '/HISMRD2/MRD/Reports/Date_wise_Surgeries/',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: null,
      method: 'GET',
      mode: 'cors',
      credentials: 'include'
    }).then((r) => r.text())

    const ControlID = res.slice(
      res.indexOf('ControlID=') + 'ControlID='.length,
      res.indexOf('&Mode')
    ).substring(0, 32)
    // downloadBlob(ControlID,'res.html','text/html')
    // window.alert( ControlID.substring(0,32))
    const orrep = await fetch(loc + '/HISMRD2/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + ControlID + '&Mode=true&OpType=Export&FileName=Get_Date_Wise_Surgeries&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML'
      , {
        headers: {
          accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
          'accept-language': 'en-US,en;q=0.9',
          'cache-control': 'no-cache',
          pragma: 'no-cache',
          'upgrade-insecure-requests': '1',
          cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
        },
        referrerPolicy: 'strict-origin-when-cross-origin',
        body: null,
        method: 'GET',
        mode: 'cors',
        credentials: 'include'
      }).then(res => res.arrayBuffer())

    writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/VBHC/ORsheets/' + br + '_' + month + '_' + year + '_OP_IP_Rep.xlsx', Buffer.from(orrep))
  }

  async function dldis(dischargeds, year, month) {
    const arr = []
    let c = 0
    for (const dis of dischargeds) {
      try {
        c++
        if (c % 50 == 0 && c > 49) {
          await page.reload()

          cookie = await page.cookies()
        }
        console.log('Downloading dissum for patinet ', br, dischargeds.indexOf(dis) + ' of ' + dischargeds.length)
        //http://130.1.2.27/HISDM/DM/Generic/DM_GetDischargeSummary?pin=1321284&ipid=310541&_=1688293975122
        const dissum = await fetch(loc + '/' + dm + '/DM/Generic/DM_GetDischargeSummary?pin=' + (+dis.PIN.split('.').pop()) + '&ipid=' + dis.IPID + '&_=1671687175418', {
          headers: {
            accept: '*/*',
            'accept-language': 'en-US,en;q=0.9',
            'x-requested-with': 'XMLHttpRequest',
            cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
          },
          referrer: loc + '/HISDM4/DM/Main',
          referrerPolicy: 'strict-origin-when-cross-origin',
          body: null,
          method: 'GET',
          mode: 'cors',
          credentials: 'include'
        }).then((response) => response.json())
        arr.push(dissum)
      } catch (error) {
        console.log(error);
      }
    }
    writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/VBHC/' + br + '_' + month + '_' + year + '_dissum.json', JSON.stringify(arr))

    // downloadBlob(JSON.stringify(arr), br + "_" + month + "_" + year + '_dissum.json', 'text/html')
  }

  async function approvalTrack() {
    await page.goto(loc + "/HISUCAF/ARApprovalMaintenance");
    await page.reload()
    cookie = await page.cookies()
    function getMonthStartAndEndDates(startDate, endDate) {
      const start = new Date(startDate);
      const end = new Date(endDate);
      const result = [];

      // Set the start date to the beginning of the month
      start.setDate(1);

      // Array to hold month abbreviations
      const monthAbbreviations = [
        'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
      ];

      // Loop through each month
      while (start <= end) {
        const month = start.getMonth();
        const year = start.getFullYear();

        // Calculate the last day of the current month
        const lastDay = new Date(year, month + 1, 0).getDate();

        // Create a new object with the start and end dates of the current month
        const startDateFormatted = formatDate(start, monthAbbreviations);
        const endDateFormatted = formatDate(new Date(year, month, lastDay), monthAbbreviations);

        result.push({
          start: startDateFormatted,
          end: endDateFormatted,
        });

        // Move to the next month
        start.setMonth(month + 1);
      }

      return result;
    }

    // Helper function to format date as DD-MMM-YYYY
    function formatDate(date, monthAbbreviations) {
      const day = String(date.getDate()).padStart(2, '0');
      const month = monthAbbreviations[date.getMonth()];
      const year = date.getFullYear();

      return `${day}-${month}-${year}`;
    }

    // Example usage
    let startDate = '2023-06-01'; // Start of the period
    let endDate = '2023-07-01'; // End of the period

    let monthDates = getMonthStartAndEndDates(startDate, endDate);
    console.log(monthDates);

    for (let month of monthDates) {
      let from = month.start
      let to = month.end

      console.log("test22 token")
      let text = await fetch(loc + "/HISUCAF/ARApprovalMaintenance", {
        "headers": {
          "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
          "accept-language": "en-US,en;q=0.9,ar;q=0.8",
          "cache-control": "max-age=0",
          "upgrade-insecure-requests": "1",
          "cookie": cookie.map(a => a.name + "=" + a.value).join("; ")
        },
        "referrerPolicy": "strict-origin-when-cross-origin",
        "body": null,
        "method": "GET",
        "mode": "cors",
        "credentials": "include"
      }).then(a => a.text())

      text = text.slice('RequestVerificationToken" type="hidden" value="'.length + text.indexOf('RequestVerificationToken" type="hidden" value=')).slice(0, 200)
      text = text.slice(0, text.indexOf('"'))
      let vertoken = text
      //console.log(vertoken,br)
      let allwithdr = await fetch(loc + "/HISUCAF/ApprovalResults/GetApprovalRequestList", {
        "headers": {
          "accept": "application/json, text/javascript, */*; q=0.01",
          "accept-language": "en-US,en;q=0.9,ar;q=0.8",
          "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
          "x-requested-with": "XMLHttpRequest",
          "cookie": cookie.map(a => a.name + "=" + a.value).join("; ")
        },
        "referrer": loc + "/HISUCAF/ApprovalResults",
        "referrerPolicy": "strict-origin-when-cross-origin",
        "body": "RegistrationNo=0&CategoryId=0&FDate=" + from + "&TDate=" + to,
        "method": "POST",
        "mode": "cors",
        "credentials": "include"
      }).then(a => a.json())
      writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Approvals/' + br + "_" + from + "_" + to + "_allwithDr.json", JSON.stringify(allwithdr.ReturnData))











      // let allrequests= await fetch(loc+"/HISUCAF/ARApprovalMaintenance/GetNaphiesApprovalRequestList_v2", {
      //   "headers": {
      //     "accept": "application/json, text/javascript, */*; q=0.01",
      //     "accept-language": "en-US,en;q=0.9,ar;q=0.8",
      //     "content-type": "multipart/form-data; boundary=----WebKitFormBoundaryl0JI6Er5fYT91B3A",
      //     "x-requested-with": "XMLHttpRequest",
      //     "cookie": cookie.map(a => a.name + "=" + a.value).join("; ")
      //   },
      //   "referrer": loc+"/HISUCAF/ARApprovalMaintenance",
      //   "referrerPolicy": "strict-origin-when-cross-origin",
      //   "body": "------WebKitFormBoundaryl0JI6Er5fYT91B3A\r\nContent-Disposition: form-data; name=\"CategoryId\"\r\n\r\n0\r\n------WebKitFormBoundaryl0JI6Er5fYT91B3A\r\nContent-Disposition: form-data; name=\"FromDate\"\r\n\r\n"+from+"\r\n------WebKitFormBoundaryl0JI6Er5fYT91B3A\r\nContent-Disposition: form-data; name=\"ToDate\"\r\n\r\n"+to+"\r\n------WebKitFormBoundaryl0JI6Er5fYT91B3A\r\nContent-Disposition: form-data; name=\"ApprovalRequestId\"\r\n\r\n0\r\n------WebKitFormBoundaryl0JI6Er5fYT91B3A\r\nContent-Disposition: form-data; name=\"__RequestVerificationToken\"\r\n\r\n"+vertoken+"\r\n------WebKitFormBoundaryl0JI6Er5fYT91B3A--\r\n",
      //   "method": "POST",
      //   "mode": "cors",
      //   "credentials": "include"
      // }).then(a=>a.json()).catch(a=>console.log(a))
      // allrequests=allrequests.ReturnData//.slice(16000)
      // writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Approvals/'+br+"_"+from+"_"+to+"_All.json",JSON.stringify(allrequests))

      //downloadBlob(JSON.stringify(allrequests),"MakkahAll.json","application/json")
      let c = 0
      let d = []
      //console.log(allrequests.length,"requests")
      const batchSize = 50;

      for (let i = 0; i < allwithdr.ReturnData.length; i += batchSize) {
        if (i % 200 == 0) {
          await page.reload()

          cookie = await page.cookies();
        }
        console.log(br, allwithdr.ReturnData.length, i, "requests")
        const currentBatch = allwithdr.ReturnData.slice(i, i + batchSize);

        try {
          const promises = currentBatch.map(req => fetch(loc + "/HISUCAF/ARApprovalMaintenance/GetApprovalRequestDetails_v2", {
            "headers": {
              "accept": "application/json, text/javascript, */*; q=0.01",
              "accept-language": "en-US,en;q=0.9,ar;q=0.8",
              "content-type": "multipart/form-data; boundary=----WebKitFormBoundaryM90rQuTVH66lHBGx",
              "x-requested-with": "XMLHttpRequest",
              "cookie": cookie.map(a => a.name + "=" + a.value).join("; ")
            },
            "referrer": loc + "/HISUCAF/ARApprovalMaintenance",
            "referrerPolicy": "strict-origin-when-cross-origin",
            "body": "------WebKitFormBoundaryM90rQuTVH66lHBGx\r\nContent-Disposition: form-data; name=\"ApprovalRequestId\"\r\n\r\n" + req.Id + "\r\n------WebKitFormBoundaryM90rQuTVH66lHBGx\r\nContent-Disposition: form-data; name=\"RequestTypeId\"\r\n\r\n1\r\n------WebKitFormBoundaryM90rQuTVH66lHBGx\r\nContent-Disposition: form-data; name=\"__RequestVerificationToken\"\r\n\r\n" + vertoken + "\r\n------WebKitFormBoundaryM90rQuTVH66lHBGx--\r\n",
            "method": "POST",
            "mode": "cors",
            "credentials": "include"
          }).then(response => response.json()).then(a => d.push(a.ReturnData))
          );
          await Promise.all(promises);
        } catch (error) {
        }
      }
      //////
      console.log(d.length, "d")
      d = d.flat()
      writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Approvals/' + br + "_" + from + "_" + to + "_Allapprovals.json", JSON.stringify(d))
      d = []
    }
    //downloadBlob(JSON.stringify(d),"syncMakkah.json","application/json")


    console.log("test tokendone ")

  }

  async function invoices(year, month) {
    const f1 = month + '_' + year
    const f2 = 'invoices'
    const f3 = br

    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1)
    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2)

    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2 + '/' + f3) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3)
    const completed = fs.readdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3).map(a => +a.split('_')[1])

    let startofmonth = new Date(year + '-' + month + '-2')
    let lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 1)
    startofmonth = startofmonth.toISOString().split('T')[0]
    lastDayOfMonth = lastDayOfMonth.toISOString().split('T')[0]

    const discharged = await fetch(loc + '/HISARADMIN/ARBillFinalization/getDischargeBills', {
      headers: {
        accept: 'application/json, text/javascript, */*; q=0.01',
        'accept-language': 'en-US,en;q=0.9',
        'cache-control': 'no-cache',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        pragma: 'no-cache',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
      },
      referrer: loc + '/HISARADMIN/ARBillFinalization',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: 'categoryId=0&fDate=01-' + month + '-' + year + '&tDate=' + lastDayOfMonth,
      method: 'POST',
      mode: 'cors',
      credentials: 'include'
    }).then((response) => response.json())
    discharged.Res = discharged.Res.filter(a => !completed.includes(+a.IPID))//.filter(a=>a.Company.includes("7001599658")||a.Company.includes("TOTAL")||a.Company.includes("TCS")||a.Company.includes("GOSI"))
    
    console.log(completed.length, 'Invoices Already downloaded', br)

    for (const a of discharged.Res) {
      if (a.Company.includes('0110') || a.Company.includes('MOH') || a.Company.includes('MOHREG') || a.Company.includes('1648')) {
        a.type = 11
      } else { a.type = 10 }
    }

    const ismains = [0, 1]

    let c = 0
    for (const invoice of discharged.Res)
    // "body": "xmldata=%5B%7B%22ipbillno%22%3A"+invoice.BillNo+"%7D%5D&rtype=7&ismulti="+type+"&invtype=1&ismain=1&rdisp=P&isdot=0",
    {
      c++
      if (c % 50 == 0 && c > 49) {
        await page.reload()

        cookie = await page.cookies()
      }
      try {


        const pdfsbuffer = []
        for (const ismain of ismains) {

          const res = await fetch(loc + '/HISARADMIN/Reports/Reports.aspx', {
            headers: {
              accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
              'accept-language': 'en-US,en;q=0.9',
              'cache-control': 'no-cache',
              'content-type': 'application/x-www-form-urlencoded',
              pragma: 'no-cache',
              'upgrade-insecure-requests': '1',
              cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
            },
            referrer: loc + '/HISARADMIN/GeneralInvoicePrinting',
            referrerPolicy: 'strict-origin-when-cross-origin',
            body: 'xmldata=%5B%7B%22ipbillno%22%3A' + invoice.BillNo + '%7D%5D&rtype=7&ismulti=' + invoice.type + '&invtype=1&ismain=' + ismain + '&rdisp=P&isdot=0',
            method: 'POST',
            mode: 'cors',
            credentials: 'include'
          }).then((response) => response.text())

          const ctrl = res.slice(
            res.indexOf('ControlID=') + 'ControlID='.length,
            res.indexOf('&Mode')
          ).substring(0, 32)
          const pdf = await fetch(loc + '/HISARADMIN/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + ctrl + '&Mode=true&OpType=Export&FileName=' + (+invoice.PIN.split('.').pop()) + '_' + invoice.BillNo + '_' + loc.split('__').pop() + '&ContentDisposition=OnlyHtmlInline&Format=PDF'

            , {
              headers: {
                accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                'accept-language': 'en-US,en;q=0.9',
                'cache-control': 'no-cache',
                pragma: 'no-cache',
                'upgrade-insecure-requests': '1',
                cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
              },
              referrerPolicy: 'strict-origin-when-cross-origin',
              body: null,
              method: 'GET',
              mode: 'cors',
              credentials: 'include'
            }).then(res => res.arrayBuffer())

          // downloadBlob(pdf,`1_${invoice.IPID}_${+invoice.PIN.split('.').pop()}_${br}_${month}_${year}_.pdf`,'application/pdf')
          pdfsbuffer.push(pdf)
          // writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/'+f1+"/"+f2+"/"+f3+`/1_${invoice.IPID}_${+invoice.PIN.split('.').pop()}_${br}_${month}_${year} (${ismain}).pdf`, Buffer.from(pdf))

          console.log(c, discharged.Res.length, "Invoices", br)
        }
        const merged = await mergePDFDocuments(pdfsbuffer)
        writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3 + `/1_${invoice.IPID}_${+invoice.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`, Buffer.from(merged))
      } catch (error) {

      }
    }
  }

  async function Adj(year, month) {
    async function getbillexcel(billNo) {
      const res = await fetch(loc + '/HISARADMIN/Reports/Reports.aspx', {
        headers: {
          accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
          'accept-language': 'en-US,en;q=0.9',
          'cache-control': 'max-age=0',
          'content-type': 'application/x-www-form-urlencoded',
          'upgrade-insecure-requests': '1',
          cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
        },
        referrerPolicy: 'strict-origin-when-cross-origin',
        body: 'billno=' + billNo + '&rtype=' + 21 + '&ismain=1',
        method: 'POST',
        mode: 'cors',
        credentials: 'include'
      }).then(a => a.text())
      const ctrl = res.slice(
        res.indexOf('ControlID=') + 'ControlID='.length,
        res.indexOf('&Mode')
      ).substring(0, 32)
      const xx = await fetch(loc + '/HISARADMIN/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + ctrl + '&Mode=true&OpType=Export&FileName=IPInvoiceMainNPD&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML', {
        headers: {
          accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
          'accept-language': 'en-US,en;q=0.9',
          'cache-control': 'no-cache',
          pragma: 'no-cache',
          'upgrade-insecure-requests': '1',
          cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
        },
        referrerPolicy: 'strict-origin-when-cross-origin',
        body: null,
        method: 'GET',
        mode: 'cors',
        credentials: 'include'
      }).then(a => a.arrayBuffer())
      return xx
    }
    const f1 = month + '_' + year
    const f2 = 'Adjust'
    const f3 = br

    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1)
    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2)

    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2 + '/' + f3) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3)
    const completed = fs.readdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3).map(a => +a.split('_')[1])

    let startofmonth = new Date(year + '-' + month + '-2')
    let lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 1)
    startofmonth = startofmonth.toISOString().split('T')[0]
    lastDayOfMonth = lastDayOfMonth.toISOString().split('T')[0]

    const discharged = await fetch(loc + '/HISARADMIN/ARBillFinalization/getDischargeBills', {
      headers: {
        accept: 'application/json, text/javascript, */*; q=0.01',
        'accept-language': 'en-US,en;q=0.9',
        'cache-control': 'no-cache',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        pragma: 'no-cache',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
      },
      referrer: loc + '/HISARADMIN/ARBillFinalization',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: 'categoryId=0&fDate=01-' + month + '-' + year + '&tDate=' + lastDayOfMonth,
      method: 'POST',
      mode: 'cors',
      credentials: 'include'
    }).then((response) => response.json())
    discharged.Res = discharged.Res.filter(a => !completed.includes(+a.IPID))
    console.log(completed.length, 'Invoices Already downloaded', br)

    for (const a of discharged.Res) {
      if (a.Company.includes('0110') || a.Company.includes('MOH') || a.Company.includes('MOHREG') || a.Company.includes('1648')) {
        a.type = 11
      } else { a.type = 10 }
    }

    const ismains = [1]

    let c = 0
    for (const invoice of discharged.Res)
    // "body": "xmldata=%5B%7B%22ipbillno%22%3A"+invoice.BillNo+"%7D%5D&rtype=7&ismulti="+type+"&invtype=1&ismain=1&rdisp=P&isdot=0",
    {
      c++
      if (c % 50 == 0 && c > 49) {
        await page.reload()

        cookie = await page.cookies()
      }
      try {



        for (const ismain of ismains) {

          const res = await fetch(loc + '/HISARADMIN/Reports/Reports.aspx', {
            headers: {
              accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
              'accept-language': 'en-US,en;q=0.9',
              'cache-control': 'no-cache',
              'content-type': 'application/x-www-form-urlencoded',
              pragma: 'no-cache',
              'upgrade-insecure-requests': '1',
              cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
            },
            referrer: loc + '/HISARADMIN/GeneralInvoicePrinting',
            referrerPolicy: 'strict-origin-when-cross-origin',
            body: 'xmldata=%5B%7B%22ipbillno%22%3A' + invoice.BillNo + '%7D%5D&rtype=7&ismulti=' + 10 + '&invtype=1&ismain=' + ismain + '&rdisp=P&isdot=0',
            method: 'POST',
            mode: 'cors',
            credentials: 'include'
          }).then((response) => response.text())

          const ctrl = res.slice(
            res.indexOf('ControlID=') + 'ControlID='.length,
            res.indexOf('&Mode')
          ).substring(0, 32)
          const pdf = await fetch(loc + '/HISARADMIN/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + ctrl + '&Mode=true&OpType=Export&FileName=' + (+invoice.PIN.split('.').pop()) + '_' + invoice.BillNo + '_' + loc.split('__').pop() + '&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML'

            , {
              headers: {
                accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                'accept-language': 'en-US,en;q=0.9',
                'cache-control': 'no-cache',
                pragma: 'no-cache',
                'upgrade-insecure-requests': '1',
                cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
              },
              referrerPolicy: 'strict-origin-when-cross-origin',
              body: null,
              method: 'GET',
              mode: 'cors',
              credentials: 'include'
            }).then(res => res.arrayBuffer())

          // downloadBlob(pdf,`1_${invoice.IPID}_${+invoice.PIN.split('.').pop()}_${br}_${month}_${year}_.pdf`,'application/pdf')
          //pdfsbuffer.push(pdf)
          writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + "/" + f2 + "/" + f3 + `/1_${invoice.IPID}_${+invoice.PIN.split('.').pop()}_${br}_${month}_${year}_after.xlsx`, Buffer.from(pdf))

          let before = await getbillexcel(invoice.BillNo)
          writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + "/" + f2 + "/" + f3 + `/1_${invoice.IPID}_${+invoice.PIN.split('.').pop()}_${br}_${month}_${year}_before.xlsx`, Buffer.from(before))


          console.log(c, discharged.Res.length, "Invoices", br)
        }
        //const merged = await mergePDFDocuments(pdfsbuffer)
        //writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3 + `/1_${invoice.IPID}_${+invoice.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`, Buffer.from(merged))
      } catch (error) {

      }
    }
  }

  async function Labs(year, month) {
    const f1 = month + '_' + year
    const f2 = 'Labs'
    const f3 = br

    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1)
    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2)
    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2 + '/' + f3) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3)
    const completed = fs.readdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3).map(a => +a.split('_')[1])

    let startofmonth = new Date(year + '-' + month + '-2')
    let lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 1)
    startofmonth = startofmonth.toISOString().split('T')[0]
    lastDayOfMonth = lastDayOfMonth.toISOString().split('T')[0]

    const discharged = await fetch(loc + '/HISARADMIN/ARBillFinalization/getDischargeBills', {
      headers: {
        accept: 'application/json, text/javascript, */*; q=0.01',
        'accept-language': 'en-US,en;q=0.9',
        'cache-control': 'no-cache',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        pragma: 'no-cache',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
      },
      referrer: loc + '/HISARADMIN/ARBillFinalization',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: 'categoryId=0&fDate=01-' + month + '-' + year + '&tDate=' + lastDayOfMonth,
      method: 'POST',
      mode: 'cors',
      credentials: 'include'
    }).then((response) => response.json())

    discharged.Res = discharged.Res.filter(a => !completed.includes(+a.IPID))
    //.filter(a=>a.Company.includes("7001599658")||a.Company.includes("TOTAL")||a.Company.includes("TCS")||a.Company.includes("GOSI"))
    console.log('already done', completed.length, br)
    let tt = 0
    let vv = 0

    for (const visit of discharged.Res) {
      try {
        vv++



        function getlablinks(html) {
          const labs = '<html>' + html + '</html'
          const dom = new JSDOM(labs)
          const { document } = dom.window

          const tblResultList = document.getElementById('tbl-result-list')

          const rows = Array.from(tblResultList.rows)

          const done = rows
            .filter((a) => a.getAttribute('data-ptype'))
            .filter((a) =>
              a.getAttribute('class').includes('green') && !(a.getAttribute('class').toLowerCase().includes('x-ray') || a.getAttribute('class').toLowerCase().includes('radio') || a.getAttribute('class').toLowerCase().includes('imaging'))
              &&
              !bl.includes(a.cells[4].textContent) &&
              new Date(a.cells[6].textContent) >= new Date(new Date(visit.AdmitDateTime) - 1000 * 60 * 60 * 24) &&
              new Date(a.cells[6].textContent) <= new Date(new Date(visit.DischargeDateTime) + 1000 * 60 * 60 * 24) &&
              !a.cells[4].textContent.includes('POCT')
            )

          const mapped = done.map((a) => {
            const orderid = a.getAttribute('data-orderid')

            const key = orderid + '_' + a.childNodes[7].textContent + '_' + a.childNodes[11].textContent + '_' + a.childNodes[13].textContent

            const testid = a.getAttribute('data-testid')
            const ptype = a.getAttribute('data-ptype').replace(1, 'True').replace(0, 'False')
            const testcomb = a.getAttribute('data-testcomb')
            //http://130.1.2.27/HISRadiology/ReportViewer/Result.aspx?isIp=false&testIds=1667&orderId=3506242
            // Link: loc + `/HISLABORATORY/AREAS/LAB/RDLFiles/LabResult.aspx?isIp=${ptype}&testids=${testcomb}&orderid=${orderid}`,

            return {
              Link2: `/HISRadiology/ReportViewer/Result.aspx?isIp=${ptype.toLowerCase()}&testIds=${testcomb.replaceAll(',', '')}&orderId=${orderid}`,

              Link: loc + `/HISLABORATORY/AREAS/LAB/RDLFiles/LabResult.aspx?isIp=${ptype}&testids=${testcomb.slice(0, -1)}&orderid=${orderid}`,
              key
            }
          }).filter((value, index, self) => self.findIndex((item) => item.key === value.key) === index)

          return mapped
        }

        const labs = await fetch(loc + '/HISPATIENTVIEW/PatientView/ResultsView/ResultsView?RegNo=' + visit.PIN.split('.').pop() + '&Panic=0&_=1660834387037', {
          headers: {
            accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'accept-language': 'en-US,en;q=0.9',
            'cache-control': 'no-cache',
            pragma: 'no-cache',
            'upgrade-insecure-requests': '1',
            cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
          },
          referrerPolicy: 'strict-origin-when-cross-origin',
          body: null,
          method: 'GET',
          mode: 'cors',
          credentials: 'include'
        }).then((response) => response.text())

        const mapped = getlablinks(labs).filter((value, index, self) => self.indexOf(value) === index)

        const key = 'Link'

        const mappedUnique = [...new Map(mapped.map(item =>
          [item[key], item])).values()]
        console.log(br, 'Found ' + mappedUnique.length + ' labs, Downloading...')
        const ctrls = []
        let pdfs = []
        let c = 0
        for (const x of mappedUnique) {
          c++
          tt++

          if (tt % 10 == 0 && tt > 10) {
            await page.reload()

            cookie = await page.cookies()
          }

          const res = await fetch(x.Link, {
            headers: {
              accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
              'accept-language': 'en-US,en;q=0.9',
              'cache-control': 'max-age=0',
              'upgrade-insecure-requests': '1',
              cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
            },
            referrerPolicy: 'strict-origin-when-cross-origin',
            body: null,
            method: 'GET',
            mode: 'cors',
            credentials: 'include'
          }).then((response) => response.text())
          // downloadBlob(res,'res.html','text/html')
          const ControlID = res.slice(
            res.indexOf('ControlID=') + 'ControlID='.length,
            res.indexOf('&Mode')
          ).substring(0, 32)

          //HISLABORATORY
          console.log(br, c + ' of ' + mappedUnique.length + ' Labs Done for patient ' + visit.PIN + ' NO ' + vv + ' of ' + discharged.Res.length)
          //http://130.1.2.27/HISRadiology/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=6f05db5f7e3748269323b88acca0298f&Mode=true&OpType=Export&FileName=XrayResult&ContentDisposition=OnlyHtmlInline&Format=PDF
          const pdf = await fetch(loc + '/HISLABORATORY/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + ControlID + '&Mode=true&OpType=Export&FileName=' + visit.PIN.split('.').pop() + '&ContentDisposition=OnlyHtmlInline&Format=PDF'
            , {
              headers: {
                accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                'accept-language': 'en-US,en;q=0.9',
                'cache-control': 'no-cache',
                pragma: 'no-cache',
                'upgrade-insecure-requests': '1',
                cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
              },
              referrerPolicy: 'strict-origin-when-cross-origin',
              body: null,
              method: 'GET',
              mode: 'cors',
              credentials: 'include'
            }).then(res => res.arrayBuffer())
          pdfs.push(pdf)
        }

        // writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/'+f1+"/"+f2+"/"+f3+`/4_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`, Buffer.from(pdfs))
        // if(pdfs.length>0)
        //  {
        let merged = await mergePDFDocuments(pdfs.reverse())
        writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3 + `/4_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`, Buffer.from(merged))
        //  }
        merged = null
        pdfs = []
      } catch (error) {
        console.log(error, br, 'Error in Labs')
      }
    }
  }
  async function LabsbyPIN(PIN) {
    const f1 = PIN
    const f2 = 'Labs'
    const f3 = br

    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1)
    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2)
    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2 + '/' + f3) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3)
    const completed = fs.readdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3).map(a => +a.split('_')[1])

    //let startofmonth = new Date(year + '-' + month + '-2')
    //let lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 1)
    //startofmonth = startofmonth.toISOString().split('T')[0]
    //lastDayOfMonth = lastDayOfMonth.toISOString().split('T')[0]

    
    let tt = 0
    let vv = 0

    
      try {
        vv++



        function getlablinks(html) {
          const labs = '<html>' + html + '</html'
          const dom = new JSDOM(labs)
          const { document } = dom.window

          const tblResultList = document.getElementById('tbl-result-list')

          const rows = Array.from(tblResultList.rows)

          const done = rows
            .filter((a) => a.getAttribute('data-ptype'))
            .filter((a) =>
              a.getAttribute('class').includes('green') 
              //&& !(a.getAttribute('class').toLowerCase().includes('x-ray') || a.getAttribute('class').toLowerCase().includes('radio') || a.getAttribute('class').toLowerCase().includes('imaging'))
              
             // !bl.includes(a.cells[4].textContent) &&
             // new Date(a.cells[6].textContent) >= new Date(new Date(visit.AdmitDateTime) - 1000 * 60 * 60 * 24) &&
             // new Date(a.cells[6].textContent) <= new Date(new Date(visit.DischargeDateTime) + 1000 * 60 * 60 * 24) &&
             // !a.cells[4].textContent.includes('POCT')
            )

          const mapped = done.map((a) => {
            const orderid = a.getAttribute('data-orderid')

            const key = orderid + '_' + a.childNodes[7].textContent + '_' + a.childNodes[11].textContent + '_' + a.childNodes[13].textContent

            const testid = a.getAttribute('data-testid')
            const ptype = a.getAttribute('data-ptype').replace(1, 'True').replace(0, 'False')
            const testcomb = a.getAttribute('data-testcomb')
            //http://130.1.2.27/HISRadiology/ReportViewer/Result.aspx?isIp=false&testIds=1667&orderId=3506242
            // Link: loc + `/HISLABORATORY/AREAS/LAB/RDLFiles/LabResult.aspx?isIp=${ptype}&testids=${testcomb}&orderid=${orderid}`,

            return {
              Link2: `/HISRadiology/ReportViewer/Result.aspx?isIp=${ptype.toLowerCase()}&testIds=${testcomb.replaceAll(',', '')}&orderId=${orderid}`,

              Link: loc + `/HISLABORATORY/AREAS/LAB/RDLFiles/LabResult.aspx?isIp=${ptype}&testids=${testcomb.slice(0, -1)}&orderid=${orderid}`,
              key
            }
          }).filter((value, index, self) => self.findIndex((item) => item.key === value.key) === index)

          return mapped
        }

        const labs = await fetch(loc + '/HISPATIENTVIEW/PatientView/ResultsView/ResultsView?RegNo=' + PIN + '&Panic=0&_=1660834387037', {
          headers: {
            accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'accept-language': 'en-US,en;q=0.9',
            'cache-control': 'no-cache',
            pragma: 'no-cache',
            'upgrade-insecure-requests': '1',
            cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
          },
          referrerPolicy: 'strict-origin-when-cross-origin',
          body: null,
          method: 'GET',
          mode: 'cors',
          credentials: 'include'
        }).then((response) => response.text())

        const mapped = getlablinks(labs).filter((value, index, self) => self.indexOf(value) === index)

        const key = 'Link'

        const mappedUnique = [...new Map(mapped.map(item =>
          [item[key], item])).values()]
        console.log(br, 'Found ' + mappedUnique.length + ' labs, Downloading...')
        const ctrls = []
        let pdfs = []
        let c = 0
        for (const x of mappedUnique) {
          c++
          tt++

          if (tt % 10 == 0 && tt > 10) {
            await page.reload()

            cookie = await page.cookies()
          }

          const res = await fetch(x.Link, {
            headers: {
              accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
              'accept-language': 'en-US,en;q=0.9',
              'cache-control': 'max-age=0',
              'upgrade-insecure-requests': '1',
              cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
            },
            referrerPolicy: 'strict-origin-when-cross-origin',
            body: null,
            method: 'GET',
            mode: 'cors',
            credentials: 'include'
          }).then((response) => response.text())
          // downloadBlob(res,'res.html','text/html')
          const ControlID = res.slice(
            res.indexOf('ControlID=') + 'ControlID='.length,
            res.indexOf('&Mode')
          ).substring(0, 32)

          //HISLABORATORY
          //console.log(br, c + ' of ' + mappedUnique.length + ' Labs Done for patient ' + visit.PIN + ' NO ' + vv + ' of ' + discharged.Res.length)
          //http://130.1.2.27/HISRadiology/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=6f05db5f7e3748269323b88acca0298f&Mode=true&OpType=Export&FileName=XrayResult&ContentDisposition=OnlyHtmlInline&Format=PDF
          const pdf = await fetch(loc + '/HISLABORATORY/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + ControlID + '&Mode=true&OpType=Export&FileName=' + PIN + '&ContentDisposition=OnlyHtmlInline&Format=PDF'
            , {
              headers: {
                accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                'accept-language': 'en-US,en;q=0.9',
                'cache-control': 'no-cache',
                pragma: 'no-cache',
                'upgrade-insecure-requests': '1',
                cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
              },
              referrerPolicy: 'strict-origin-when-cross-origin',
              body: null,
              method: 'GET',
              mode: 'cors',
              credentials: 'include'
            }).then(res => res.arrayBuffer())
          pdfs.push(pdf)
        }

        // writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/'+f1+"/"+f2+"/"+f3+`/4_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`, Buffer.from(pdfs))
        // if(pdfs.length>0)
        //  {
        let merged = await mergePDFDocuments(pdfs.reverse())
        writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3 + `/4_${PIN}.pdf`, Buffer.from(merged))
        //  }
        merged = null
        pdfs = []
      } catch (error) {
        console.log(error, br, 'Error in Labs')
      }
    
  }
  async function Rads(year, month) {
    const f1 = month + '_' + year
    const f2 = 'Rads'
    const f3 = br

    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1)
    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2)
    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2 + '/' + f3) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3)
    const completed = fs.readdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3).map(a => +a.split('_')[1])

    let startofmonth = new Date(year + '-' + month + '-2')
    let lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 1)
    startofmonth = startofmonth.toISOString().split('T')[0]
    lastDayOfMonth = lastDayOfMonth.toISOString().split('T')[0]

    const discharged = await fetch(loc + '/HISARADMIN/ARBillFinalization/getDischargeBills', {
      headers: {
        accept: 'application/json, text/javascript, */*; q=0.01',
        'accept-language': 'en-US,en;q=0.9',
        'cache-control': 'no-cache',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        pragma: 'no-cache',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
      },
      referrer: loc + '/HISARADMIN/ARBillFinalization',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: 'categoryId=0&fDate=01-' + month + '-' + year + '&tDate=' + lastDayOfMonth,
      method: 'POST',
      mode: 'cors',
      credentials: 'include'
    }).then((response) => response.json())

    discharged.Res = discharged.Res.filter(a => !completed.includes(+a.IPID))//.filter(a=>a.Company.includes("7001599658")||a.Company.includes("TOTAL")||a.Company.includes("TCS")||a.Company.includes("GOSI"))
    
    console.log('already done', "Rads", completed.length, br)
    let tt = 0
    let vv = 0
    for (const visit of discharged.Res) {
      try {




        function getlablinks(html) {
          const labs = '<html>' + html + '</html'
          const dom = new JSDOM(labs)
          const { document } = dom.window

          const tblResultList = document.getElementById('tbl-result-list')

          const rows = Array.from(tblResultList.rows)
          //console.log(done[0])
          const done = rows
            .filter((a) => a.getAttribute('data-ptype'))
            .filter((a) =>
              a.getAttribute('class').includes('green') &&
              (a.getAttribute('class').toLowerCase().includes('x-ray') || a.getAttribute('class').toLowerCase().includes('radio') || a.getAttribute('class').toLowerCase().includes('imaging'))
              &&
              !bl.includes(a.cells[4].textContent) &&
              new Date(a.cells[6].textContent) >= new Date(new Date(visit.AdmitDateTime) - 1000 * 60 * 60 * 24) &&
              new Date(a.cells[6].textContent) <= new Date(new Date(visit.DischargeDateTime) + 1000 * 60 * 60 * 24) &&
              !a.cells[4].textContent.includes('POCT')
            )

          const mapped = done.map((a) => {
            const orderid = a.getAttribute('data-orderid')

            //const key = orderid + '_' + a.childNodes[7].textContent + '_' + a.childNodes[11].textContent + '_' + a.childNodes[13].textContent

            const testid = a.getAttribute('data-testid')
            const ptype = a.getAttribute('data-ptype').replace(1, 'True').replace(0, 'False')
            const testcomb = a.getAttribute('data-testcomb')
            //http://130.1.2.27/HISRadiology/ReportViewer/Result.aspx?isIp=false&testIds=1667&orderId=3506242
            // Link: loc + `/HISLABORATORY/AREAS/LAB/RDLFiles/LabResult.aspx?isIp=${ptype}&testids=${testcomb}&orderid=${orderid}`,

            return {
              Link2: loc + `/HISRadiology/ReportViewer/Result.aspx?isIp=${ptype}&testIds=${testid.replaceAll(',', '')}&orderId=${orderid}`,

              Link: loc + `/HISLABORATORY/AREAS/LAB/RDLFiles/LabResult.aspx?isIp=${ptype}&testids=${testcomb}&orderid=${orderid}`,
              //key
            }
          })//.filter((value, index, self) => self.findIndex((item) => item.key === value.key) === index)

          return mapped
        }

        const labs = await fetch(loc + '/HISPATIENTVIEW/PatientView/ResultsView/ResultsView?RegNo=' + visit.PIN.split('.').pop() + '&Panic=0&_=1660834387037', {
          headers: {
            accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'accept-language': 'en-US,en;q=0.9',
            'cache-control': 'no-cache',
            pragma: 'no-cache',
            'upgrade-insecure-requests': '1',
            cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
          },
          referrerPolicy: 'strict-origin-when-cross-origin',
          body: null,
          method: 'GET',
          mode: 'cors',
          credentials: 'include'
        }).then((response) => response.text())

        const mapped = getlablinks(labs)//.filter((value, index, self) => self.indexOf(value) === index)

        const key = 'Link'

        console.log(br, 'Found ' + mapped.length + ' Rads, Downloading...')
        const ctrls = []
        let pdfs = []
        let c = 0

        for (const x of mapped) {
          c++
          tt++

          if (tt % 100 == 0 && tt > 99) {
            await page.reload()

            cookie = await page.cookies()
          }

          const res = await fetch(x.Link2, {
            headers: {
              accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
              'accept-language': 'en-US,en;q=0.9',
              'cache-control': 'max-age=0',
              'upgrade-insecure-requests': '1',
              cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
            },
            referrerPolicy: 'strict-origin-when-cross-origin',
            body: null,
            method: 'GET',
            mode: 'cors',
            credentials: 'include'
          }).then((response) => response.text())
          // downloadBlob(res,'res.html','text/html')
          const ControlID = res.slice(
            res.indexOf('ControlID=') + 'ControlID='.length,
            res.indexOf('&Mode')
          ).substring(0, 32)
          //console.log(ControlID,x.Link2)
          //HISLABORATORY
          console.log(br, c + ' of ' + mapped.length + ' Rads Done for patient ' + visit.PIN + ' NO ' + vv + ' of ' + discharged.Res.length)
          //http://130.1.2.27/HISRadiology/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=6f05db5f7e3748269323b88acca0298f&Mode=true&OpType=Export&FileName=XrayResult&ContentDisposition=OnlyHtmlInline&Format=PDF
          let pdf = await fetch(loc + '/HISRadiology/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + ControlID + '&Mode=true&OpType=Export&FileName=' + visit.PIN.split('.').pop() + '&ContentDisposition=OnlyHtmlInline&Format=PDF'
            , {
              headers: {
                accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                'accept-language': 'en-US,en;q=0.9',
                'cache-control': 'no-cache',
                pragma: 'no-cache',
                'upgrade-insecure-requests': '1',
                cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
              },
              referrerPolicy: 'strict-origin-when-cross-origin',
              body: null,
              method: 'GET',
              mode: 'cors',
              credentials: 'include'
            }).then(res => res.arrayBuffer())
          //writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3 + `/4_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}_${c}.pdf`, Buffer.from(pdf))

          //pdf=new Uint8Array(pdf)
          pdfs.push(pdf)
        }

        //writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/'+f1+"/"+f2+"/"+f3+`/4_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`, Buffer.from(pdfs))
        if (true || pdfs.length > 0) {
          let merged = await mergePDFDocuments(pdfs.reverse())
          writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3 + `/4_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`, Buffer.from(merged))
          merged = null
        }

        pdfs = []
      } catch (error) {
        console.log(error, br, 'Error in Labs')
      }
    }
  }
  async function CS(year, month) {
    const f1 = month + '_' + year
    const f2 = 'CS'
    const f3 = br

    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1)
    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2)
    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2 + '/' + f3) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3)
    const completed = fs.readdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3).map(a => +a.split('_')[1])

    let startofmonth = new Date(year + '-' + month + '-2')
    let lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 1)
    startofmonth = startofmonth.toISOString().split('T')[0]
    lastDayOfMonth = lastDayOfMonth.toISOString().split('T')[0]

    const discharged = await fetch(loc + '/HISARADMIN/ARBillFinalization/getDischargeBills', {
      headers: {
        accept: 'application/json, text/javascript, */*; q=0.01',
        'accept-language': 'en-US,en;q=0.9',
        'cache-control': 'no-cache',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        pragma: 'no-cache',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
      },
      referrer: loc + '/HISARADMIN/ARBillFinalization',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: 'categoryId=0&fDate=01-' + month + '-' + year + '&tDate=' + lastDayOfMonth,
      method: 'POST',
      mode: 'cors',
      credentials: 'include'
    }).then((response) => response.json())

    discharged.Res = discharged.Res.filter(a => (a.Company.includes('0110') || a.Company.includes('MOH') || a.Company.includes('MOHREG') || a.Company.includes('1648'))
      && !completed.includes(+a.IPID))
    console.log('already done', "CS", completed.length, br)
    let tt = 0

    for (const visit of discharged.Res) {
      try {


        tt++

        if (tt % 100 == 0 && tt > 99) {
          await page.reload()

          cookie = await page.cookies()
        }

        function getlablinks(html) {
          const labs = '<html>' + html + '</html'
          const dom = new JSDOM(labs)
          const { document } = dom.window

          const tblResultList = document.getElementById('tbl-result-list')

          const rows = Array.from(tblResultList.rows)
          //console.log(done[0])
          const done = rows
            .filter((a) => a.getAttribute('data-ptype'))
            .filter((a) =>
              a.getAttribute('class').includes('green') &&



              new Date(a.cells[6].textContent) >= new Date(new Date(visit.AdmitDateTime) - 1000 * 60 * 60 * 24) &&
              new Date(a.cells[6].textContent) <= new Date(new Date(visit.DischargeDateTime) + 1000 * 60 * 60 * 24) &&
              a.cells[5].textContent.toLowerCase().includes('c/s')
            )

          const mapped = done.map((a) => {
            const orderid = a.getAttribute('data-orderid')

            //const key = orderid + '_' + a.childNodes[7].textContent + '_' + a.childNodes[11].textContent + '_' + a.childNodes[13].textContent

            const testid = a.getAttribute('data-testid')
            const ptype = a.getAttribute('data-ptype').replace(1, 'True').replace(0, 'False')
            const testcomb = a.getAttribute('data-testcomb')
            //http://130.1.2.27/HISRadiology/ReportViewer/Result.aspx?isIp=false&testIds=1667&orderId=3506242
            // Link: loc + `/HISLABORATORY/AREAS/LAB/RDLFiles/LabResult.aspx?isIp=${ptype}&testids=${testcomb}&orderid=${orderid}`,

            return {
              Link2: loc + `/HISLABORATORY/AREAS/LAB/RDLFiles/LabResult.aspx?isIp=${ptype}&testids=${testid.replaceAll(',', '')}&orderid=${orderid}`,

              Link: loc + `/HISLABORATORY/AREAS/LAB/RDLFiles/LabResult.aspx?isIp=${ptype}&testids=${testcomb}&orderid=${orderid}`,
              //key
            }
          })//.filter((value, index, self) => self.findIndex((item) => item.key === value.key) === index)

          return mapped
        }

        const labs = await fetch(loc + '/HISPATIENTVIEW/PatientView/ResultsView/ResultsView?RegNo=' + visit.PIN.split('.').pop() + '&Panic=0&_=1660834387037', {
          headers: {
            accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'accept-language': 'en-US,en;q=0.9',
            'cache-control': 'no-cache',
            pragma: 'no-cache',
            'upgrade-insecure-requests': '1',
            cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
          },
          referrerPolicy: 'strict-origin-when-cross-origin',
          body: null,
          method: 'GET',
          mode: 'cors',
          credentials: 'include'
        }).then((response) => response.text())

        const mapped = getlablinks(labs)//.filter((value, index, self) => self.indexOf(value) === index)

        const key = 'Link'

        mapped.length > 0 ? console.log(br, 'Found ' + mapped.length + ' CS, Downloading...') : null
        const ctrls = []
        let pdfs = []
        let c = 0

        for (const x of mapped) {
          c++

          const res = await fetch(x.Link2, {
            headers: {
              accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
              'accept-language': 'en-US,en;q=0.9',
              'cache-control': 'max-age=0',
              'upgrade-insecure-requests': '1',
              cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
            },
            referrerPolicy: 'strict-origin-when-cross-origin',
            body: null,
            method: 'GET',
            mode: 'cors',
            credentials: 'include'
          }).then((response) => response.text())
          // downloadBlob(res,'res.html','text/html')
          const ControlID = res.slice(
            res.indexOf('ControlID=') + 'ControlID='.length,
            res.indexOf('&Mode')
          ).substring(0, 32)
          //console.log(ControlID,x.Link2)
          //HISLABORATORY
          console.log(br, c + ' of ' + mapped.length + ' CS Done for patient ' + visit.PIN + ' NO ' + tt + ' of ' + discharged.Res.length)
          //http://130.1.2.27/HISRadiology/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=6f05db5f7e3748269323b88acca0298f&Mode=true&OpType=Export&FileName=XrayResult&ContentDisposition=OnlyHtmlInline&Format=PDF
          let pdf = await fetch(loc + '/HISLABORATORY/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + ControlID + '&Mode=true&OpType=Export&FileName=' + visit.PIN.split('.').pop() + '&ContentDisposition=OnlyHtmlInline&Format=PDF'
            , {
              headers: {
                accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                'accept-language': 'en-US,en;q=0.9',
                'cache-control': 'no-cache',
                pragma: 'no-cache',
                'upgrade-insecure-requests': '1',
                cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
              },
              referrerPolicy: 'strict-origin-when-cross-origin',
              body: null,
              method: 'GET',
              mode: 'cors',
              credentials: 'include'
            }).then(res => res.arrayBuffer())
          //writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3 + `/4_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}_${c}.pdf`, Buffer.from(pdf))

          //pdf=new Uint8Array(pdf)
          pdfs.push(pdf)
        }

        //writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/'+f1+"/"+f2+"/"+f3+`/4_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`, Buffer.from(pdfs))
        if (true || pdfs.length > 0) {
          let merged = await mergePDFDocuments(pdfs.reverse())
          writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3 + `/4_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`, Buffer.from(merged))
          merged = null
        }

        pdfs = []
      } catch (error) {
        console.log(error, br, 'Error in CS')
      }
    }
  }








  async function DischargePDF(year, month) {
    const f1 = month + '_' + year
    const f2 = 'DischargeSumm'
    const f3 = br

    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1)
    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2)
    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2 + '/' + f3) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3)
    const completed = fs.readdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3).map(a => +a.split('_')[1])

    let startofmonth = new Date(year + '-' + month + '-2')
    let lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 1)
    startofmonth = startofmonth.toISOString().split('T')[0]
    lastDayOfMonth = lastDayOfMonth.toISOString().split('T')[0]

    const discharged = await fetch(loc + '/HISARADMIN/ARBillFinalization/getDischargeBills', {
      headers: {
        accept: 'application/json, text/javascript, */*; q=0.01',
        'accept-language': 'en-US,en;q=0.9',
        'cache-control': 'no-cache',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        pragma: 'no-cache',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
      },
      referrer: loc + '/HISARADMIN/ARBillFinalization',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: 'categoryId=0&fDate=01-' + month + '-' + year + '&tDate=' + lastDayOfMonth,
      method: 'POST',
      mode: 'cors',
      credentials: 'include'
    }).then((response) => response.json())

    discharged.Res = discharged.Res.filter(a => !completed.includes(+a.IPID))//.filter(a=>a.Company.includes("7001599658")||a.Company.includes("TOTAL")||a.Company.includes("TCS")||a.Company.includes("GOSI"))
    
    console.log('already done', completed.length, br)
    let tt = 0

    for (const visit of discharged.Res) {
      tt++
      console.log(br, tt, discharged.Res.length, 'discharge summaries');
      if (tt % 100 == 0 && tt > 99) {
        await page.reload()

        cookie = await page.cookies()
      }
      const dm = br == 'Jeddah' ? 'HISDM' : 'HISDM4'
      let pdf = await fetch(loc + "/" + dm + "/Areas/DM/Report/DischargeSummary.aspx?ipid=" + visit.IPID + "&vid=0", {
        "headers": {
          "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
          "accept-language": "en-US,en;q=0.9",
          "cache-control": "no-cache",
          "pragma": "no-cache",
          "upgrade-insecure-requests": "1"
        },
        "referrerPolicy": "strict-origin-when-cross-origin",
        "body": null,
        "method": "GET",
        "mode": "cors",
        "credentials": "include"
      }).then(res => res.arrayBuffer())



      writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3 + `/2_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`, Buffer.from(pdf))
      //  }

    }
  }
  async function progressnotesPDF(year, month) {
    const f1 = month + '_' + year
    const f2 = 'ProgressNotes'
    const f3 = br

    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1)
    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2)
    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 & '/' + f2 + '/' + f3) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3)
    const completed = fs.readdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3).map(a => +a.split('_')[1])

    let startofmonth = new Date(year + '-' + month + '-2')
    let lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 1)
    startofmonth = startofmonth.toISOString().split('T')[0]
    lastDayOfMonth = lastDayOfMonth.toISOString().split('T')[0]

    const discharged = await fetch(loc + '/HISARADMIN/ARBillFinalization/getDischargeBills', {
      headers: {
        accept: 'application/json, text/javascript, */*; q=0.01',
        'accept-language': 'en-US,en;q=0.9',
        'cache-control': 'no-cache',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        pragma: 'no-cache',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
      },
      referrer: loc + '/HISARADMIN/ARBillFinalization',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: 'categoryId=0&fDate=01-' + month + '-' + year + '&tDate=' + lastDayOfMonth,
      method: 'POST',
      mode: 'cors',
      credentials: 'include'
    }).then((response) => response.json())

    discharged.Res = discharged.Res.filter(a => !completed.includes(+a.IPID))//.filter(a=>a.Company.includes("7001599658")||a.Company.includes("TOTAL")||a.Company.includes("TCS")||a.Company.includes("GOSI"))
    
    console.log('already done', completed.length, br)
    let tt = 0

    for (const visit of discharged.Res) {
      tt++
      console.log(br, tt, discharged.Res.length, 'ProgressNotes');
      if (tt % 30 == 0 && tt > 30) {
        await page.reload()

        cookie = await page.cookies()
      }
      const dm = br == 'Jeddah' ? 'HISDM' : 'HISDM4'
      let res = await fetch(loc + "/HISMRD2/Reports/MRDReportViewer.aspx?rpt=106&VisitID=" + visit.IPID + "&RegNo=" + visit.PIN.split('.').pop() + "&IP_OP=IP", {
        "headers": {
          "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
          "accept-language": "en-US,en;q=0.9,ar;q=0.8",
          "upgrade-insecure-requests": "1",
          cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
        },
        "referrer": loc + "/HISMRD2/MRD/File/Patient_Folder/",
        "referrerPolicy": "strict-origin-when-cross-origin",
        "body": null,
        "method": "GET",
        "mode": "cors",
        "credentials": "include"
      }).then(a => a.text())
      const ctrl = res.slice(
        res.indexOf('ControlID=') + 'ControlID='.length,
        res.indexOf('&Mode')
      ).substring(0, 32)

      //http://130.2.10.21/HISMRD2/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID="+ctrlid+"&Mode=true&OpType=Export&FileName=ProgressNote&ContentDisposition=OnlyHtmlInline&Format=PDF
      const pdf = await fetch(loc + "/HISMRD2/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=" + ctrl + "&Mode=true&OpType=Export&FileName=ProgressNote&ContentDisposition=OnlyHtmlInline&Format=PDF"

        , {
          headers: {
            accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'accept-language': 'en-US,en;q=0.9',
            'cache-control': 'no-cache',
            pragma: 'no-cache',
            'upgrade-insecure-requests': '1',
            cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
          },
          referrerPolicy: 'strict-origin-when-cross-origin',
          body: null,
          method: 'GET',
          mode: 'cors',
          credentials: 'include'
        }).then(res => res.arrayBuffer())


      writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3 + `/3_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`, Buffer.from(pdf))
      //  }

    }
  }


  async function getdistext(visit) {
    let findpt = await fetch(loc + "/SGH-MRD/MEDREP/Medical/FindPatient?id=" + (+visit.PIN.split('.')[1]) + "&_=1687682413235", {
      "headers": {
        "accept": "*/*",
        "accept-language": "en-US,en;q=0.9,ar;q=0.8",
        "x-requested-with": "XMLHttpRequest",
        cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
      },
      "referrer": loc + "/SGH-MRD/MEDREP/Medical",
      "referrerPolicy": "strict-origin-when-cross-origin",
      "body": null,
      "method": "GET",
      "mode": "cors",
      "credentials": "include"
    }).then(a => a.json())
    findpt = findpt.find(a => a.IPID == visit.IPID)
    let x = await fetch(loc + "/SGH-MRD/MEDREP/Medical/GetReport?id=" + findpt.ROWID + "&_=1687682413237", {
      "headers": {
        "accept": "*/*",
        "accept-language": "en-US,en;q=0.9,ar;q=0.8",
        "x-requested-with": "XMLHttpRequest",
        cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
      },
      "referrer": loc + "/SGH-MRD/MEDREP/Medical",
      "referrerPolicy": "strict-origin-when-cross-origin",
      "body": null,
      "method": "GET",
      "mode": "cors",
      "credentials": "include"
    }).then(a => a.json())
    //console.log(findpt)

    let t = ""
    var textreport = x.report;
    var dt = "";
    var rPTName = x.PTName;
    var rPin = x.pin
    var rAge = x.Age
    var rSex = x.Sex;
    var rNation = x.Nationality
    t += '<html><head>'
    t += '<style type="text/css" media="print">'
    t += '@page '
    t += '{'
    t += 'size: auto;'
    t += 'margin-top: 250px;margin-left:20px;margin-right:20px;th:text-align:left;'
    t += '}'
    t += '#ad{ display:none;}'
    t += '#leftbar{ display:none;}'
    t += '#contentarea{ width:100%;}'
    t += 'body '
    t += '{'
    t += 'background-color:#FFFFFF; '
    t += 'border: none;font-size:16px;'
    t += '} th{text-align:left;font-size:16px;} </style></head><body><div>'
    t += '<p style="font-size:16px;font-family:Calibri;">' + dt + '</p><br/>'
    t += '<pre style="top:20px;font-family:Calibri;font-size:16px;white-space: pre-wrap;font-weight:500;">'
    t+='<img id="img-header" src='+loc+'/SGH-MRD/Images/mreportheader.png width="800"></img>'
    //t += '<img id="img-header" src=' + branch.logo + ' width="800"></img>'
    t += x.ReportHeader + '<br>'
    t += textreport.replace(/\n/gi, '<br>')
    t += x.ReportFoot

    t += '</pre>'

    t += '</div></body></html>'

    return t
  }

  async function getORtext(visit) {//http://130.3.2.208/SGH-MRD/MEDREP/operation/FindPatient?id=2313084&_=1691298421153

    let findpts = await fetch(loc + "/SGH-MRD/MEDREP/operation/FindPatient?id=" + (+visit.PIN.split('.')[1]) + "&_=1687682413235", {
      "headers": {
        "accept": "*/*",
        "accept-language": "en-US,en;q=0.9,ar;q=0.8",
        "x-requested-with": "XMLHttpRequest",
        cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
      },
      "referrer": loc + "/SGH-MRD/MEDREP/Medical",
      "referrerPolicy": "strict-origin-when-cross-origin",
      "body": null,
      "method": "GET",
      "mode": "cors",
      "credentials": "include"
    }).then(a => a.json())
    //console.log(findpts.length,findpts.map(a=>a.OPR_DATE))
    findpts = findpts.filter(a => new Date(a.OPR_DATE_ORDER) >= new Date(visit.AdmitDateTime).addDays(-1) && new Date(a.OPR_DATE_ORDER) <= new Date(visit.DischargeDateTime).addDays(1))
    //console.log(findpts.length)
    let ors = []
    //http://130.3.2.208/SGH-MRD/MEDREP/operation/ViewPatDet?id=2313084&row=78022&_=1691298421154
    for (let findpt of findpts) {

      let x = await fetch(loc + "/SGH-MRD/MEDREP/operation/ViewPatDet?id=" + findpt.PIN + "&row=" + findpt.ROWID + "&_=1687682413237", {
        "headers": {
          "accept": "*/*",
          "accept-language": "en-US,en;q=0.9,ar;q=0.8",
          "x-requested-with": "XMLHttpRequest",
          cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
        },
        "referrer": loc + "/SGH-MRD/MEDREP/Medical",
        "referrerPolicy": "strict-origin-when-cross-origin",
        "body": null,
        "method": "GET",
        "mode": "cors",
        "credentials": "include"
      }).then(a => a.json())
      //console.log(findpt)

      let t = ""

      // t+='<img id="img-header" src='+loc+'/SGH-MRD/Images/mreportheader.png width="800"></img>'
      t += '<html><head>'

      t += '<style type="text/css" media="print">'
      t += '@page '
      t += '{'
      t += 'size: auto;'
      /* t+='margin-top: 72px;');*/
      t += '}'
      t += 'body '
      t += '{'
      t += 'background-color:#FFFFFF; '
      t += 'border: none;'
      t += 'margin-left: 0px;margin-right: 0px;'/*margin-top: 72px;*/

      t += '}</style></head><body><div>'
      t+='<img id="img-header" src='+loc+'/SGH-MRD/Images/mreportheader.png width="800"></img>'
      //t += '<img id="img-header" src=' + branch.logo + ' width="800"></img>'
      t += '<table style="font-size:12px;font-family:Courier New;padding:0px;">'
      t += '<tr><td style="width: 150px;">Patient Name</td><td style="padding-right: 10px;">:</td><td style="width: 200px;">' + x.PT_NAME + '</td></tr>'
      t += '<tr><td>PIN Number</td><td>:</td><td>' + x.PIN + '</td></tr>'
      t += '</table>'

      t += '<table style="font-size:12px;font-family:Courier New;">'
      t += '<tr><td style="width: 20px;">Age</td><td>:</td><td style="width: 100px;padding:0px;">' + x.Age + '</td><td style="width:50px;">Sex</td><td style="padding-right: 10px;">:</td>'
      t += '<td style="width: 120px;">' + x.SEX + '</td><td style="width: 100px; padding-left: 20px; padding-right: 20px;">Room Number</td><td style="padding-right: 10px;">:</td>'
      t += '<td>' + x.room_no + '</td></tr></table>'
      t += '<table style="font-size:12px;font-family:Courier New;">'
      t += '<tr><td style="width: 150px;">Operated Date</td><td>:</td><td style="width: 70px;">' + new Date(x.OPR_DATE).addDays(1).toUTCString().split(", ")[1].slice(0, 11) + '</td><td style="width: 120px; padding-left: 10px;">Nationality</td>'
      t += '<td style="padding-right: 10px;">:</td><td>' + x.NATIONALITY + '</td><td colspan="2"></td></tr>'
      t += '</table>'

      var doc1 = x.SURG_CODE + " - " + x.doc_name;
      var doc2 = x.assistant_code1 + " - " + x.assistant_name1
      var doc3 = x.assistant_code2 + " - " + x.assistant_name2
      var ancode = x.ana_code + " - " + x.ANANAME;
      var orCode = x.OR_CODE;


      t += '<table style="font-size:12px;font-family:Courier New;padding:0px;">'
      t += '<tr><td style="width:110px;">SURGEON</td><td style="padding-right: 2px;">:</td><td style="width:300px;">' + doc1 + '</td><td style="width:150px;padding-left:1px;">ANAESTHESIOLOGIST</td><td style="width:5px;padding-left: 2px;">:</td><td style="width:50px;padding-left:5px;">' + ancode + '</td></tr>'
      t += '<tr><td>1st ASSISTANT</td><td style="padding-right: 2px;">:</td><td style="width:300px;">' + doc2 + '</td><td style="width:150px;padding-left:1px;">ANAESTHESIA</td><td style="width:5px;padding-left: 2px;">:</td><td style="width:50px;padding-left:5px;">' + x.ana_type + '</td></tr>'
      t += '<tr><td>2nd ASSISTANT</td style="padding-right: 2px;"><td>:</td><td style="width:300px;">' + doc3 + '</td><td style="width:150px;padding-left:1px;">SCRUB NURSE</td><td style="width:5px;padding-left: 2px;">:</td><td style="width:50px;padding-left:5px;">' + x.scrub_nurse + '</td></tr>'
      t += '<tr><td>OPERATION BEGAN</td><td style="padding-right: 2px;">:</td><td style="width:300px;">' + x.opr_time + '</td><td style="width:150px;padding-left:1px;">ANAESTHESIA BEGAN</td><td style="width:5px;padding-left: 2px;">:</td><td style="width:50px;padding-left:5px;">' + x.ana_begun + '</td></tr>'
      t += '<tr><td>OPERATION ENDED</td><td style="padding-right: 2px;">:</td><td style="width:300px;">' + x.opr_end + '</td><td style="width:150px;padding-left:1px;">ANAESTHESIA END</td><td style="width:5px;padding-left: 2px;">:</td><td style="width:50px;padding-left:5px;">' + x.ana_end + '</td></tr>'
      t += '<tr><td>OR Code</td><td style="padding-right: 2px;">:</td><td style="width:300px;">' + orCode + '</td><td style="width:150px;padding-left:1px;">ANAESTHESIA END</td><td style="width:5px;padding-left: 2px;">:</td><td style="width:50px;padding-left:5px;">' + x.ana_end + '</td></tr>'
      t += '</table>'

      t += '<br />'
      t += '<p style="font-family:Courier New;font-size:12px;">'
      t += x.report.replace(/\n/gi, '<br>')
      t += '</p>'
      t += '</div></body></html>'
      ors.push(t)

    }

    return ors
  }
  async function MedicalReport(year, month) {
    try {


      const f1 = month + '_' + year
      const f2 = 'MedicalReport'
      const f3 = br

      fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1)
      fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2)
      fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2 + '/' + f3) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3)
      const completed = fs.readdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3).map(a => +a.split('_')[1])

      let startofmonth = new Date(year + '-' + month + '-2')
      let lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 1)
      startofmonth = startofmonth.toISOString().split('T')[0]
      lastDayOfMonth = lastDayOfMonth.toISOString().split('T')[0]

      const discharged = await fetch(loc + '/HISARADMIN/ARBillFinalization/getDischargeBills', {
        headers: {
          accept: 'application/json, text/javascript, */*; q=0.01',
          'accept-language': 'en-US,en;q=0.9',
          'cache-control': 'no-cache',
          'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
          pragma: 'no-cache',
          'x-requested-with': 'XMLHttpRequest',
          cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
        },
        referrer: loc + '/HISARADMIN/ARBillFinalization',
        referrerPolicy: 'strict-origin-when-cross-origin',
        body: 'categoryId=0&fDate=01-' + month + '-' + year + '&tDate=' + lastDayOfMonth,
        method: 'POST',
        mode: 'cors',
        credentials: 'include'
      }).then((response) => response.json())

      discharged.Res = discharged.Res.filter(a => !completed.includes(+a.IPID))//.filter(a=>a.Company.includes("7001599658")||a.Company.includes("TOTAL")||a.Company.includes("TCS")||a.Company.includes("GOSI"))
    
      console.log('already done', completed.length, br)
      let tt = 0
      let errs = []
      const recpage = await browser.newPage();
      //try {


      for (const visit of discharged.Res) {

        tt++
        try {
          console.log(br, tt, discharged.Res.length, 'MedicalReport');
          if (tt % 100 == 0 && tt > 99) {
            await page.reload()

            cookie = await page.cookies()
          }
          let reporttext = await getdistext(visit)
          await recpage.setContent(reporttext, { waitUntil: 'domcontentloaded' });

          // To reflect CSS used for screens instead of print
          await recpage.emulateMediaType('screen');

          // Downlaod the PDF
          await recpage.pdf({
            path: 'C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3 + `/3_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`,
            format: "A4"
          });
          //writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3 + `/3_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`, Buffer.from(pdf))
          //  }
        } catch (error) {
          errs.push(visit)

        }
      }

      writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3 + `/${br}_${month}_${year}_errors.json`, JSON.stringify(errs))
    } catch (error) {
      console.log(error, br, month, year, 'MedicalReport')

    }
  }














  async function ORReport(year, month) {
    try {


      const f1 = month + '_' + year
      const f2 = 'OR'
      const f3 = br

      fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1)
      fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2)
      fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2 + '/' + f3) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3)
      const completed = fs.readdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3).map(a => +a.split('_')[1])

      let startofmonth = new Date(year + '-' + month + '-2')
      let lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 1)
      startofmonth = startofmonth.toISOString().split('T')[0]
      lastDayOfMonth = lastDayOfMonth.toISOString().split('T')[0]

      const discharged = await fetch(loc + '/HISARADMIN/ARBillFinalization/getDischargeBills', {
        headers: {
          accept: 'application/json, text/javascript, */*; q=0.01',
          'accept-language': 'en-US,en;q=0.9',
          'cache-control': 'no-cache',
          'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
          pragma: 'no-cache',
          'x-requested-with': 'XMLHttpRequest',
          cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
        },
        referrer: loc + '/HISARADMIN/ARBillFinalization',
        referrerPolicy: 'strict-origin-when-cross-origin',
        body: 'categoryId=0&fDate=01-' + month + '-' + year + '&tDate=' + lastDayOfMonth,
        method: 'POST',
        mode: 'cors',
        credentials: 'include'
      }).then((response) => response.json())

      discharged.Res = discharged.Res.filter(a => !completed.includes(+a.IPID))//.filter(a=>a.Company.includes("7001599658")||a.Company.includes("TOTAL")||a.Company.includes("TCS")||a.Company.includes("GOSI"))
    
      console.log('already done', completed.length, br)
      let tt = 0
      let errs = []
      const recpage = await browser.newPage();
      //try {


      for (const visit of discharged.Res) {

        tt++
        try {
          console.log(br, tt, discharged.Res.length, 'ORS');
          if (tt % 100 == 0 && tt > 99) {
            await page.reload()

            cookie = await page.cookies()
          }
          let reporttexts = await getORtext(visit)

          let x = 0
          if (reporttexts.length > 0) {
            for (let reporttext of reporttexts) {
              x++
              await recpage.setContent(reporttext, { waitUntil: 'domcontentloaded' });

              // To reflect CSS used for screens instead of print
              await recpage.emulateMediaType('screen');

              // Downlaod the PDF
              await recpage.pdf({
                path: 'C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3 + `/3_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}_${x}.pdf`,
                format: "A4"
              })
            }
          }
          //writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3 + `/3_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`, Buffer.from(pdf))
          //  }
        } catch (error) {
          errs.push(visit)

        }
      }

      writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3 + `/${br}_${month}_${year}_errors.json`, JSON.stringify(errs))
    } catch (error) {
      console.log(error, br, month, year, 'MedicalReport')

    }
  }
  async function mergeAndReport(year, month) {
    const f1 = month + '_' + year
    const f2 = 'Merged'
    const f3 = br

    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1)
    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2)
    fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3) ? console.log('already there', 'C:/Users/mis1.ryd/' + f1 + '/' + f2 + '/' + f3) : fs.mkdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3)
    const completed = fs.readdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3).map(a => +a.split('_')[1])

    let startofmonth = new Date(year + '-' + month + '-2')
    let lastDayOfMonth = new Date(startofmonth.getFullYear(), startofmonth.getMonth() + 1, 1)
    startofmonth = startofmonth.toISOString().split('T')[0]
    lastDayOfMonth = lastDayOfMonth.toISOString().split('T')[0]

    const discharged = await fetch(loc + '/HISARADMIN/ARBillFinalization/getDischargeBills', {
      headers: {
        accept: 'application/json, text/javascript, */*; q=0.01',
        'accept-language': 'en-US,en;q=0.9',
        'cache-control': 'no-cache',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        pragma: 'no-cache',
        'x-requested-with': 'XMLHttpRequest',
        cookie: cookie.map(a => a.name + '=' + a.value).join('; ')
      },
      referrer: loc + '/HISARADMIN/ARBillFinalization',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: 'categoryId=0&fDate=01-' + month + '-' + year + '&tDate=' + lastDayOfMonth,
      method: 'POST',
      mode: 'cors',
      credentials: 'include'
    }).then((response) => response.json())

    discharged.Res


    let c = 0
    for (const visit of discharged.Res//.filter(a=>a.Company.includes("7001599658")||a.Company.includes("TOTAL")||a.Company.includes("TCS")||a.Company.includes("GOSI"))
    ) {
      try {


        c++
        console.log(br, c, discharged.Res.length, 'merge');
        let merge = []
        visit.Labs = fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "Labs" + '/' + f3 + `/4_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`
        ) && fs.statSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "Labs" + '/' + f3 + `/4_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`
        ).size > 15000
        visit.Rads = fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "Rads" + '/' + f3 + `/4_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`
        ) && fs.statSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "Rads" + '/' + f3 + `/4_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`
        ).size > 15000
        visit.Culture = fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "CS" + '/' + f3 + `/4_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`
        ) && fs.statSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "CS" + '/' + f3 + `/4_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`
        ).size > 15000
        visit.invoice = fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "invoices" + '/' + f3 + `/1_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`
        )&&fs.statSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "invoices" + '/' + f3 + `/1_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`
        ).size>270000
        visit.progressnotes = fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "ProgressNotes" + '/' + f3 + `/3_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`
        )
        visit.MedicalReport = fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "MedicalReport" + '/' + f3 + `/3_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`
        )
        visit.DischargeSumm = fs.existsSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "DischargeSumm" + '/' + f3 + `/2_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`
        ) && fs.statSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "DischargeSumm" + '/' + f3 + `/2_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`
        ).size > 30000
        visit.ORS = fs.readdirSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "OR" + '/' + f3).filter(a => a.includes(`3_${visit.IPID}_${+visit.PIN.split('.').pop()}`))
        console.log(visit.ORS.length)
        visit.ORS = visit.ORS.length > 0 ? visit.ORS.map(a => fs.readFileSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "OR" + '/' + f3 + `/` + a)) : null


        visit.invoice ? merge.push(fs.readFileSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "invoices" + '/' + f3 + `/1_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`)) : null

        visit.DischargeSumm ? merge.push(fs.readFileSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "DischargeSumm" + '/' + f3 + `/2_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`)) : null
        visit.ORS && visit.ORS.length > 0 ? merge.push(...visit.ORS) : null
        
        visit.MedicalReport ? merge.push(fs.readFileSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "MedicalReport" + '/' + f3 + `/3_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`)) : null
        visit.progressnotes ? merge.push(fs.readFileSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "ProgressNotes" + '/' + f3 + `/3_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`)) : null

        visit.Rads ? merge.push(fs.readFileSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "Rads" + '/' + f3 + `/4_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`)) : null
        visit.Culture ? merge.push(fs.readFileSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "CS" + '/' + f3 + `/4_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`)) : null
        visit.Labs ? merge.push(fs.readFileSync('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + "Labs" + '/' + f3 + `/4_${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`)) : null

        visit.merge = true
        const merged = await mergePDFDocuments(merge.flat())
       

        writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + '/' + f2 + '/' + f3 + `/${visit.IPID}_${+visit.PIN.split('.').pop()}_${br}_${month}_${year}.pdf`, Buffer.from(merged))
        visit.ORS && visit.ORS.length > 0 ?visit.ORS=true : false
      } catch (error) {
        visit.merge = false
        console.log(error)

      }
    }

    writeFile('C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/' + f1 + `/${br}_${month}_${year}_Summ.csv`, convertJSONtoCSV(discharged.Res))

  }
  async function censusIPD(year, month) {
    console.log(month, year, 'CENSUSIPD')

    const DTF = moment('01 ' + month + ' ' + year).format('MM/DD/YYYY')
    const DTT = moment(getlastday(month) + ' ' + month + ' ' + year).format('MM/DD/YYYY')
    /// HISMRD2/Reports/MRDReportViewer.aspx?rpt=30&FromDate="+x.start+"%20&ToDate=%20"+x.end+"&Mode=2&NationalityID=0&DepartmentId=0&DoctorID=0&SexID=0&ICD=0&CompanyID=0&CodeType=3
    const res = await fetch(loc + '/HISMRD2/Reports/MRDReportViewer.aspx?rpt=30&FromDate=' + DTF + '%20&ToDate=%20' + DTT + '&Mode=2&NationalityID=0&DepartmentId=0&DoctorID=0&SexID=0&ICD=0&CompanyID=0&CodeType=3', {
      headers: {
        accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'accept-language': 'en-US,en;q=0.9',
        'upgrade-insecure-requests': '1',
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
      },
      referrer: 'http://130.1.2.27/HISMRD2/MRD/Reports/OP_Patients_ICD_Code/',
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: null,
      method: 'GET',
      mode: 'cors',
      credentials: 'include'
    }).then(a => a.text())
    const ctrl = res.slice(
      res.indexOf('ControlID=') + 'ControlID='.length,
      res.indexOf('&Mode')
    ).substring(0, 32)
    const link = loc + '/HISMRD2/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=' + ctrl + '&Mode=true&OpType=Export&FileName=OP_Patients_ICD_Code&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML'
    const pdf = await fetch(link, {
      headers: {
        accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'accept-language': 'en-US,en;q=0.9',
        'cache-control': 'no-cache',
        pragma: 'no-cache',
        'upgrade-insecure-requests': '1',
        cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')

      },
      referrerPolicy: 'strict-origin-when-cross-origin',
      body: null,
      method: 'GET',
      mode: 'cors',
      credentials: 'include'
    }).then(a => a.arrayBuffer())
    writeFile(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/DailyCensus/${br}_${DTT.replaceAll('/', '_')}_dailyCensusIP.xlsx`, Buffer.from(pdf))

    return pdf
  }






  async function getopdVBHC(month, year) {
    let folder = `C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/VBHC/OPD/${br}/${month}_${year}/`

    let foldersync = fs.readdirSync(folder)

    if (!foldersync.includes(`${br}_${month}_${year}_ICDS.csv`)) {
      const listofBills = await fetch(loc + "/HISFRONTOFFICE/PatientVisits/GetVisitDetails", {
        "headers": {
          "accept": "application/json, text/javascript, */*; q=0.01",
          "accept-language": "en-US,en;q=0.9,ar;q=0.8",
          "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
          "x-requested-with": "XMLHttpRequest",
          cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
        },
        "referrer": loc + "/HISFRONTOFFICE/PatientVisits",
        "referrerPolicy": "strict-origin-when-cross-origin",
        "body": "docid=0&fdate=01-" + month + "-" + year + "&tdate=" + getlastday(month) + "-" + month + "-" + year,
        "method": "POST",
        "mode": "cors",
        "credentials": "include"
      }).then(a => a.json())
      writeFile(folder + `${br}_${month}_${year}_listofBills.json`, JSON.stringify(listofBills.Res))
      const csvv = converter.json2csvAsync(listofBills.Res.flat(), options)

      writeFile(folder + `${br}_${month}_${year}_listofBills.csv`, await csvv)
      let pins2 = listofBills.Res.map(a => a.PIN).filter((value, index, array) => array.indexOf(value) === index)
      let pins = listofBills.Res.map(a => a.PIN).filter((value, index, array) => array.indexOf(value) === index);
      let downloaded = foldersync.filter(a => a.includes('.json') && (!a.includes('.csv') && !a.includes('err') && !a.includes('ICD') && !a.includes('list'))).map(a => a.split('_')[0])

      pins = pins.filter(a => !downloaded.includes(a))
      console.log("downloaded", downloaded.length, "remaining", pins.length)
      const batchSize = 50
      let tt = 0
      let d = []
      let err = []

      for (let i = 0; i < pins.length; i += batchSize) {
        tt++
        if (tt % 20 == 0 && tt > 19) {
          await page.reload()

          cookie = await page.cookies()
        }
        console.log("processing", month, year, br, pins.length, i, "requests")
        const currentBatch = pins.slice(i, i + batchSize);

        try {

          const promises = currentBatch.map(req => fetch(loc + "/" + dm + "/DM/Patient/PatientFolderListAct?pin=" + (+req.split(".")[1]) + "&_=1689059294585", {
            "headers": {
              "accept": "*/*",
              "accept-language": "en-US,en;q=0.9,ar;q=0.8",
              "x-requested-with": "XMLHttpRequest",
              cookie: cookie.map(a => a.name + '=' + a.value + ';Expires=-1').join('; ')
            },
            "referrer": "http://130.2.10.21/HISDM4/DM/Main",
            "referrerPolicy": "strict-origin-when-cross-origin",
            "body": null,
            "method": "GET",
            "mode": "cors",
            "credentials": "include"
          }).then(response => response.json()).then(a => writeFile(folder + `${req}_${br}_${month}_${year}.json`, JSON.stringify({ Pin: req, Res: a.x.filter(x => true || x.attr2 == 'Chief Complaints' || x.attr2 == 'Diagnosis') }))));

          //.then(response=>response.json()).then(a=>d.push( {Pin:req,Res:a.x.filter(x=>x.attr2=='Chief Complaints'||x.attr2=='Diagnosis')})));
          await Promise.all(promises);
        } catch (error) {
          console.log(error)
          err.push(currentBatch)

        }
      }

      let allfiles = fs.readdirSync(folder).filter(a => a.includes('.json') && (!a.includes('.csv') && !a.includes('ICD') && !a.includes('err') && !a.includes('list')))
      let xt = 0
      if (allfiles.length + 100 >= pins2.length) {
        for (let file of allfiles) {
          xt++
          if (xt % 100 == 0) { console.log(allfiles.length, xt, br, month, year, "merge") }
          let d1 = JSON.parse(fs.readFileSync(folder + file))
          d1.Res = d1.Res.filter(x => x.attr2 == 'Chief Complaints' || x.attr2 == 'Diagnosis')
          //console.log(d1,file,allfiles.length,xt,br,month,year,"merge")
          d.push(d1)
        }

        //console.log(d)
        d = d.flat()
        d = d.map(a => a.Res.map(x => ({ Pin: a.Pin, ...x }))).flat()
        writeFile(folder + `${br}_${month}_${year}_ICDS.json`, JSON.stringify(d.flat()))
        const csvv2 = converter.json2csvAsync(d.flat(), options)

        writeFile(folder + `${br}_${month}_${year}_ICDS.csv`, await csvv2)
        writeFile(folder + `${br}_${month}_${year}_errs.json`, JSON.stringify(err.flat()))
        console.log("deleteing", month, year, br)
        for (let file of allfiles) {
          fs.unlinkSync(folder + file)
        }

      }


    }
    console.log("done", month, year, br)
  }

  async function VBHC() {
    if (branch.type == 'HCP') {
      for (const month of months) {
        for (const year of years) {
          const dischargeds = await getDis(year, month)
          await progressnotes(dischargeds, year, month)
          await dateWiseSurgey(year, month)
          await ORIPRep(year, month)
          await dldis(dischargeds, year, month)
        }
      }
    }
  }

  async function main() {
    for (const month of months) {
      for (const year of years) {
        await getdisSheet(year, month)

      }
    }
    //     /////////////BE
    console.log('filter', br, 'dissheet')
    if (branch.type == 'HCP') {
      for (const month of months) {
        for (const year of years) {
          await getDis(year, month)
          await BillingEff(year, month)
          await censusIPD(year, month)
        }
      }
    }
     console.log('filter', br, 'BE')
    await ipclaims()
    console.log('filter', br, 'ipclaims')
    await opclaims()
    console.log('filter', br, 'opclaims')
    await IPD3()
    console.log('filter', br, 'IPCASH')
    await PAM()
    console.log('filter', br, 'PAM')
    await opdcash()
    console.log('filter', br, 'opdcash')
    
    ////await censusOPD('2023-09-22', '2023-09-30')
    console.log('filter', br, 'censusOPD')

    console.log('filter', br, 'censusIPD')
   await Census('02-Feb-2024', new Date().toUTCString().split(', ')[1].slice(0, 11).replaceAll(' ', '-'))
    console.log('filter', br, 'CensusCOmp')
    await CensusDr()
    console.log('filter', br, 'CensusDr')
  }



  async function printer() {
    if (branch.type == "HCP") {
      if (branch.hasMR) {
        await MedicalReport('2023', 'Oct')
        await ORReport('2023','Oct')
      }
      await Labs("2023", "Oct")
      
      //await invoices("2023", "Sep")
      await progressnotesPDF("2023", "Oct")
      await CS("2023", "Oct")
      await Rads("2023", "Oct")
      await DischargePDF('2023', 'Oct')
      
      //await mergeAndReport('2023', 'Sep')
      //await mergeAndReport('2023', 'Apr')


      //await mergeAndReport('2023','Jul')
    }

  }
  async function mapSBS(){
    if (branch.isnph)
    {
    await page.goto(loc+"/HISRCMS")
    cookie = await page.cookies()
    let textv = await fetch(loc + "/HISRCMS", {
      "headers": {
        "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "accept-language": "en-US,en;q=0.9,ar;q=0.8",
        "cache-control": "max-age=0",
        "upgrade-insecure-requests": "1",
        "cookie": cookie.map(a => a.name + "=" + a.value).join("; ")
      },
      "referrerPolicy": "strict-origin-when-cross-origin",
      "body": null,
      "method": "GET",
      "mode": "cors",
      "credentials": "include"
    }).then(a => a.text())
  
    textv = textv.slice('__RequestVerificationToken" type="hidden" value="'.length + textv.indexOf('__RequestVerificationToken" type="hidden" value=')).slice(0, 200)
    textv = textv.slice(0, textv.indexOf('"'))
    let vertokenv = textv
    console.log(vertokenv.length,"vertokenv")
  for (let x of sbsmapping.filter(a=>
    //false&&
    a.Branch==branch.Name))
   { 
      console.log(x)
      let vv= await fetch(loc+"/HISRCMS/SBSMappingProfileDetails/ProcSBSDetail", {
      "headers": {
        "accept": "application/json, text/javascript, */*; q=0.01",
        "accept-language": "en-US,en;q=0.9,ar;q=0.8",
        "content-type": "multipart/form-data; boundary=----WebKitFormBoundaryhD96eU25qURZofei",
        "x-requested-with": "XMLHttpRequest",
        "cookie": cookie.map(a => a.name + "=" + a.value).join("; ")
      },
      "referrer": loc+"/HISRCMS/SBSMappingProfileDetails",
      "referrerPolicy": "strict-origin-when-cross-origin",
      "body": "------WebKitFormBoundaryhD96eU25qURZofei\r\nContent-Disposition: form-data; name=\"__RequestVerificationToken\"\r\n\r\n"+vertokenv+"\r\n------WebKitFormBoundaryhD96eU25qURZofei\r\nContent-Disposition: form-data; name=\"Id\"\r\n\r\n"+x.Id+"\r\n------WebKitFormBoundaryhD96eU25qURZofei\r\nContent-Disposition: form-data; name=\"IntegrationItemMappingProfileId\"\r\n\r\n"+x.IntegrationItemMappingProfileId+"\r\n------WebKitFormBoundaryhD96eU25qURZofei\r\nContent-Disposition: form-data; name=\"InternalCode\"\r\n\r\n"+x.InternalCode+"\r\n------WebKitFormBoundaryhD96eU25qURZofei\r\nContent-Disposition: form-data; name=\"InternalDescription\"\r\n\r\n"+x.InternalDescription+"\r\n------WebKitFormBoundaryhD96eU25qURZofei\r\nContent-Disposition: form-data; name=\"AgreedCode\"\r\n\r\n"+x.AgreedCode+"\r\n------WebKitFormBoundaryhD96eU25qURZofei\r\nContent-Disposition: form-data; name=\"AgreedCodeHyphen\"\r\n\r\n"+x.AgreedCodeHyphen+"\r\n------WebKitFormBoundaryhD96eU25qURZofei\r\nContent-Disposition: form-data; name=\"AgreedDescription\"\r\n\r\n"+x.AgreedDescription+"\r\n------WebKitFormBoundaryhD96eU25qURZofei\r\nContent-Disposition: form-data; name=\"CodeSystem\"\r\n\r\n"+x.CodeSystem+"\r\n------WebKitFormBoundaryhD96eU25qURZofei\r\nContent-Disposition: form-data; name=\"CodeSystemURL\"\r\n\r\n"+x.CodeSystemURL+"\r\n------WebKitFormBoundaryhD96eU25qURZofei\r\nContent-Disposition: form-data; name=\"GTIN\"\r\n\r\n"+x.GTIN+"\r\n------WebKitFormBoundaryhD96eU25qURZofei\r\nContent-Disposition: form-data; name=\"Deleted\"\r\n\r\nfalse\r\n------WebKitFormBoundaryhD96eU25qURZofei--\r\n",
      "method": "POST",
      "mode": "cors",
      "credentials": "include"
    }).then(a=>a.text())
  console.log(vv)
  }
     
    
      //const f = await page.$("[name='__RequestVerificationToken']")
      //console.log(f)
      //const text = await (await f.getProperty('textContent')).jsonValue()
      //console.log(text)
      let pl=await getNphiesPl()


   // console.log(pl)
  
      let nphies = converter.json2csvAsync(pl.flat(), options);
      writeFile(
        `C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Maps/${br}_nphiesMapping.csv`,
        await nphies
     );
     fs.writeFileSync(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Maps/${br}_nphiesMapping.json`, JSON.stringify(pl.flat()))

      
    //console.log(prctmapping)



  }}


  async function mapPRCT(){
    if (branch.isnph)
    {
    await page.goto(loc+"/HISRCMS")
    cookie = await page.cookies()
    let textv = await fetch(loc + "/HISRCMS", {
      "headers": {
        "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "accept-language": "en-US,en;q=0.9,ar;q=0.8",
        "cache-control": "max-age=0",
        "upgrade-insecure-requests": "1",
        "cookie": cookie.map(a => a.name + "=" + a.value).join("; ")
      },
      "referrerPolicy": "strict-origin-when-cross-origin",
      "body": null,
      "method": "GET",
      "mode": "cors",
      "credentials": "include"
    }).then(a => a.text())
  
    textv = textv.slice('__RequestVerificationToken" type="hidden" value="'.length + textv.indexOf('__RequestVerificationToken" type="hidden" value=')).slice(0, 200)
    textv = textv.slice(0, textv.indexOf('"'))
    let vertokenv = textv
    console.log(vertokenv.length,"vertokenv")

  for (let x of prctmissing.filter(a=>
    //false&&
    a.Branch==branch.Name) )
   { 
      console.log(x)
      let vv= await fetch(loc+"/HISRCMS/PractitionerMapping/SaveAndUpdatePractitionerMapping", {
        "headers": {
          "accept": "application/json, text/javascript, */*; q=0.01",
          "accept-language": "en-US,en;q=0.9,ar;q=0.8",
          "content-type": "multipart/form-data; boundary=----WebKitFormBoundaryepsk2k2wVQwTCRnx",
          "x-requested-with": "XMLHttpRequest",
          "cookie": cookie.map(a => a.name + "=" + a.value).join("; ")
        },
        "referrer": loc+"/HISRCMS/PractitionerMapping",
        "referrerPolicy": "strict-origin-when-cross-origin",
        "body": "------WebKitFormBoundaryepsk2k2wVQwTCRnx\r\nContent-Disposition: form-data; name=\"PractitionerMappingId\"\r\n\r\n0\r\n------WebKitFormBoundaryepsk2k2wVQwTCRnx\r\nContent-Disposition: form-data; name=\"DoctorId\"\r\n\r\n"+x.id+"\r\n------WebKitFormBoundaryepsk2k2wVQwTCRnx\r\nContent-Disposition: form-data; name=\"RoleId\"\r\n\r\n1\r\n------WebKitFormBoundaryepsk2k2wVQwTCRnx\r\nContent-Disposition: form-data; name=\"SpeicialtyId\"\r\n\r\n"+x.SpeicialtyId+"\r\n------WebKitFormBoundaryepsk2k2wVQwTCRnx\r\nContent-Disposition: form-data; name=\"DepartmentId\"\r\n\r\n"+x.DeptId+"\r\n------WebKitFormBoundaryepsk2k2wVQwTCRnx\r\nContent-Disposition: form-data; name=\"LicenseId\"\r\n\r\n"+x.id+"\r\n------WebKitFormBoundaryepsk2k2wVQwTCRnx\r\nContent-Disposition: form-data; name=\"Deleted\"\r\n\r\nfalse\r\n------WebKitFormBoundaryepsk2k2wVQwTCRnx\r\nContent-Disposition: form-data; name=\"__RequestVerificationToken\"\r\n\r\n"+vertokenv+"\r\n------WebKitFormBoundaryepsk2k2wVQwTCRnx--\r\n",
        "method": "POST",
        "mode": "cors",
        "credentials": "include"
      }).then(a=>a.text())
  console.log(vv)
  }
     
     

     let prctmapping=await fetch(loc+"/HISRCMS/PractitionerMapping/GetPractitionerMapping", {
      "headers": {
        "accept": "application/json, text/javascript, */*; q=0.01",
        "accept-language": "en-US,en;q=0.9,ar;q=0.8",
        "content-type": "multipart/form-data; boundary=----WebKitFormBoundaryCCbqFXW55fzoE4a6",
        "x-requested-with": "XMLHttpRequest",
        "cookie": cookie.map(a => a.name + "=" + a.value).join("; ")
      },
      "referrer": loc+"/HISRCMS/PractitionerMapping",
      "referrerPolicy": "strict-origin-when-cross-origin",
      "body": "------WebKitFormBoundaryCCbqFXW55fzoE4a6\r\nContent-Disposition: form-data; name=\"__RequestVerificationToken\"\r\n\r\n"+vertokenv+"\r\n------WebKitFormBoundaryCCbqFXW55fzoE4a6--\r\n",
      "method": "POST",
      "mode": "cors",
      "credentials": "include"
    }).then(a=>a.json())
    fs.writeFileSync(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Maps/${br}_prctmapping.json`, JSON.stringify(prctmapping))
     



  }}
     

async function mapMOH()
{ 
  let MOHMapping= await fetch(loc+"/HISITADMIN/ITADMIN/MOHCodesMapping/MohJedRiyadMastertable?deptid=0&serviceid=0&_=1698745238123", {
    "headers": {
      "accept": "*/*",
      "accept-language": "en-US,en;q=0.9,ar;q=0.8",
      "x-requested-with": "XMLHttpRequest",
      "cookie": cookie.map(a => a.name + "=" + a.value).join("; ")
    },
    "referrer": loc+"/HISITADMIN/ITADMIN/MOHCodesMapping",
    "referrerPolicy": "strict-origin-when-cross-origin",
    "body": null,
    "method": "GET",
    "mode": "cors",
    "credentials": "include"
  }).then(a=>a.json())
    fs.writeFileSync(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Maps/${br}_MOHmapping.json`, JSON.stringify(MOHMapping))



}



async function compProcedMap()
{
  async function compmap(cat,pt,id,name)
  {
let x= await fetch(loc+"/HISARADMIN/CompanyProcedureMapping/GetCompanyProcedureMapping", {
"headers": {
  "accept": "application/json, text/javascript, */*; q=0.01",
  "accept-language": "en-US,en;q=0.9,ar;q=0.8",
  "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
  "x-requested-with": "XMLHttpRequest",
  "cookie": cookie.map(a => a.name + "=" + a.value).join("; ")
},
"referrer": loc+"/HISARADMIN/CompanyProcedureMapping",
"referrerPolicy": "strict-origin-when-cross-origin",
"body": "CategoryId="+cat+"&PType="+pt+"&ServiceId="+id+"&CompanyId=0",
"method": "POST",
"mode": "cors",
"credentials": "include"
}).then(a=>a.json()).then(a=>a.Res)
      return x

      
  }

let allcomps= await fetch(loc+"/HISARADMIN/Common/get_common_list", {
"headers": {
  "accept": "application/json, text/javascript, */*; q=0.01",
  "accept-language": "en-US,en;q=0.9,ar;q=0.8",
  "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
  "x-requested-with": "XMLHttpRequest",
  "cookie": cookie.map(a => a.name + "=" + a.value).join("; ")
},
"referrer": loc+"/HISARADMIN/CompanyProcedureMapping",
"referrerPolicy": "strict-origin-when-cross-origin",
"body": "id=0&ctype=-200",
"method": "POST",
"mode": "cors",
"credentials": "include"
}).then(a=>a.json()).then(a=>a.CL)
//downloadBlob(JSON.stringify(allcomps), br+"_comps.json",'text/json')  
fs.writeFileSync(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Maps/CompMap/${br}_comps.json`, JSON.stringify(allcomps))
let CategoryIds=allcomps//.filter(a=>a.Name.toUpperCase().includes("TAWU"))
let v=0
for (let CategoryId of CategoryIds)
{
  v++
console.log(v,CategoryIds.length,br,CategoryId,"Comp map")
  

var opdclist=await 
  fetch(loc+"/HISARADMIN/Common/get_common_list", {
"headers": {
  "accept": "application/json, text/javascript, */*; q=0.01",
  "accept-language": "en-US,en;q=0.9,ar;q=0.8",
  "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
  "x-requested-with": "XMLHttpRequest",
  "cookie": cookie.map(a => a.name + "=" + a.value).join("; ")
},
"referrer": loc+"/HISARADMIN/CompanyProcedureMapping",
"referrerPolicy": "strict-origin-when-cross-origin",
"body": "id=0&ctype=60",
"method": "POST",
"mode": "cors",
"credentials": "include"
}).then(a=>a.json()).then(a=>a.CL).then(a=>a.map(x=>({PType:2,CategoryId:CategoryId.Id,...x})))
      let m=[]

for (let x of opdclist)
  {
      try {
let cm=await compmap(x.CategoryId,x.PType,x.Id,x.Name.trim())
     let xx= cm.map(a=>({
          CategoryId:x.CategoryId,
          Name:x.Name.trim(),
         ServiceId:x.Id,
              PType:x.PType,...a
              
  })

          
      )
        
      m.push(xx)
} catch (error) {
  
}
      
  }





var ipdlist= await fetch(loc+"/HISARADMIN/Common/get_common_list", {
"headers": {
  "accept": "application/json, text/javascript, */*; q=0.01",
  "accept-language": "en-US,en;q=0.9,ar;q=0.8",
  "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
  "x-requested-with": "XMLHttpRequest",
  "cookie": cookie.map(a => a.name + "=" + a.value).join("; ")
},
"referrer": loc+"/HISARADMIN/CompanyProcedureMapping",
"referrerPolicy": "strict-origin-when-cross-origin",
"body": "id=0&ctype=26",
"method": "POST",
"mode": "cors",
"credentials": "include"
}).then(a=>a.json()).then(a=>a.CL).then(a=>a.map(x=>({PType:1,CategoryId:CategoryId.Id,...x})))

for (let x of ipdlist)
  { try {
    
let cm=await compmap(x.CategoryId,x.PType,x.Id,x.Name.trim())
     let xx= cm.map(a=>({
          CategoryId:x.CategoryId, ServiceId:x.Id,
         Name:x.Name.trim(),
              PType:x.PType,...a
              
  })

          
      )
      m.push(xx)
    } catch (error) {
  
}

      
  }
let plfinal=m.flat()
fs.writeFileSync(`C:/Users/mis1.ryd/OneDrive - Saudi German Hospital (1)/Maps/CompMap/${br}_${CategoryId.Id}_cat.json`, JSON.stringify(plfinal))
//downloadBlob(JSON.stringify(plfinal), br+"_"+CategoryId.Id+".json",'text/json')  


}
}
console.log("starting comp map",br)
//await compProcedMap()
console.log("Done comp map",br)
await mapMOH()
console.log("MOH map",br)

await mapSBS()
     await  mapPRCT()

 //await printer()
  //if (branch.type == "HCP") {
  //await mergeAndReport('2023', 'Sep')}
   
  //await censusOPD('2023-09-22', '2023-11-13')
 
  
   await DoctorSchedule("01+Feb+2024","29+FEb+2024")
   
   await main()
  console.log('filter', br, 'CensusDr')
  /// ///////////////////
  // 
    

  page.close()
  console.log(br, 'Done')
  promises.push(branch)
  async function fetle() {
    if (promises.length == branchs.length) {
      await browser.close()
      throw new Error('Done all')
    }
  }
  const handle = setInterval(fetle, 30000)
}
)
