// Presented by BrilliantPy

// Editable
let templateSlideId = '1bo0DIMyJWabW2o-zejBqRtigMP8Zb4gZk1t_m1wok3s';
let folderResponsePdfId = '1cPNu9FSHg3fQN6nYf7Y1eqGLU9Re-D_J';
let folderResponseSlideId = '1YoRVpj2e4skjLMAj1nmiMQeOfgJUuJzy';
let sheetName = 'การตอบแบบฟอร์ม 1';
let pdf_file_name = "register_";
let send_status_col = 'U';
let data_begin_row = 2;
// let email_send = []; // set default no email to send
let email_send = ['nattysung101@gmail.com']; // set default one  email to send
// let email_send = ['brilliantpy1.live@gmail.com','brilliantpy2.live@gmail.com']; // set default multi email to send
var email_subject = 'ขอบคุณสำหรับ';
var email_message = 'แบบฟอร์มการสมัครได้จัดส่งให้ท่านแล้ว';

let index_col = {'ประทับเวลา':0,'เลขเคลม':1,'วันที่บันทึก':2,'เพศ':3,'ชื่อ-สกุล':4,'อายุ':5,'อีเมล์':6,'ทะเบียน':7,'หมวดจังหวัด':8,'ยี่ห้อ':9,'สี':10,'วันที่เกิดเหตุ':11,'เวลาเกิดเหตุ':12,'สถานที่เกิดเหตุ':13,'ถ้อยคำให้การ':14,'ลายเซ็นผู้ให้ถ้อยคำ':15,'เพศ[นาย]':16,'เพศ[นางสาว]':17,'เพศ[นาง]':18,'ที่อยู่':19,'send_status':20};

let colAllImage = [
  {[index_col['ลายเซ็นผู้ให้ถ้อยคำ']] : '{{ลายเซ็นผู้ให้ถ้อยคำ}}'},
];
// let colAllImage = [
//   {[index_col['รูปโปรไฟล์1']] : '{{รูปโปรไฟล์1}}'},
//   {[index_col['รูปโปรไฟล์2']] : '{{รูปโปรไฟล์2}}'},
// ];

let index_col_checkbox = [
  { [index_col['เพศ']] : [{'นาย':'R'},{'นางสาว':'S'},{'นาง':'T'}] },
  ];
// let index_col_checkbox = [
//   { [index_col['เพศ']] : [{'ชาย':'G'},{'หญิง':'H'}] },
//   { [index_col['สถานะ']] : [{'โสด':'I'},{'แต่งงาน':'J'},{'หม้าย':'K'},{'แยกกัน':'L'}] },
//   ];

// ############################################################################################################### //

// Init
let newSlideName = 'New_FormToSlidePDF_';
let sent_status = 'SENT';
let ss,sheet,lastRow,lastCol,range,values;
let data_name;
let newSlide,newSlideId,presentation,all_shape;
let titleName;
let exportPdf,pdf_name_full;
let _email_send = email_send;

function formToSlidePdf() {
  initSpreadSheet().then(async function() {
    formatTitle();
    for (let i = data_begin_row; i <= lastRow; i++) {
      clearVal();
      let cur_data = values[i-1];
      data_name = cur_data[index_col['ชื่อ-สกุล']];
      let _status = cur_data[index_col['send_status']];
      if (_status == sent_status) {
        continue;
      }
      await duplicateSlide().then(async function() {
        await updateCheckboxCol(cur_data,i).then(async function() {
          values = range.getValues();
          cur_data = values[i-1];
          await updateSlideData(cur_data).then(async function() {
            presentation.saveAndClose();
            await createPdf().then(async function() {
              let cur_email = cur_data[index_col['อีเมล์']];
              if (cur_email !== '') {
                _email_send.push(cur_email);
              }
              console.log(_email_send);
              for (let j = 0; j < _email_send.length; j++) {
                if (validateEmail(_email_send[j])) {
                  await sendEmailWithAttachment(_email_send[j]).then(function() {
                    if (j == _email_send.length-1) {
                      removeTempSlide();
                      updateStatusSent(i);
                    }
                  });
                }
              }
            });
          });
        });
      });
    }
    console.log('Program completed');
  });  
}

function clearVal() {
  data_name = '';
  newSlide = newSlideId = presentation = '';
  exportPdf = pdf_name_full = '';
  console.log('clearVal completed');
  _email_send = email_send;
  all_shape = '';
}

async function initSpreadSheet() {
  return new Promise(function(resolve) {
    ss = SpreadsheetApp.getActive();
    sheet = ss.getSheetByName(sheetName);
    lastRow = sheet.getLastRow();
    lastCol = sheet.getLastColumn();
    range = sheet.getDataRange();
    values = range.getValues();
    resolve();
    console.log('initSpreadSheet completed');
  });
}

function formatTitle() {
  titleName = values[0];
  titleName.forEach(function (item, index) {
    titleName[index] = '{{'+item+'}}';
  });
  console.log('formatTitle completed');
}

async function duplicateSlide() {
  return new Promise(function(resolve) {
    let templateSlide = DriveApp.getFileById(templateSlideId);
    let templateResponseFolder = DriveApp.getFolderById(folderResponseSlideId);
    newSlide = templateSlide.makeCopy(newSlideName.concat(data_name) , templateResponseFolder);
    resolve();
    console.log('duplicateSlide completed');
  });
}

async function updateCheckboxCol(cur_data,numRow) {
  return new Promise(function(resolve) {  
    index_col_checkbox.forEach(function (item) {
      Object.keys(item).forEach(function(key) {
        var cur_checkbox_val = cur_data[key];
        item[key].forEach(function (item_ele) {
          Object.keys(item_ele).forEach(function(key_item_ele) {
            if (key_item_ele === cur_checkbox_val) {
              sheet.getRange(item_ele[key_item_ele].concat(numRow)).setValue('✓');
            }
          })
        })
      });
    });
    resolve();
    console.log('updateCheckboxCol completed');
  });
}

async function updateSlideData(cur_data) {
  return new Promise(function(resolve) {
    // Init
    newSlideId = newSlide.getId();
    presentation = SlidesApp.openById(newSlideId);
    let slide = presentation.getSlides()[0];
    all_shape = slide.getShapes();
    titleName.forEach(async function (item, index) {
      colAllImage.forEach(async function (img_item) {
        Object.keys(img_item).forEach(async function(key) {
          if (item === img_item[key]) {
            all_shape.forEach(async function(s) {
            if (s.getText().asString().includes(img_item[key])) {
              let cur_img_url = cur_data[key];
              let new_url = formatUrlImg(cur_img_url);
              await replaceImage(s,new_url).then(async function() {
                await console.log('replace image')
              });
            }
          });
            presentation.replaceWithImage(new_url)
          }
        });
      })
      let templateVariable = item;
      let replaceValue = cur_data[index];
      presentation.replaceAllText(templateVariable, replaceValue);
  })
    resolve();
    console.log('updateSlideData completed');
  });
}
async function replaceImage(s,new_url) {
  let res;
  return new Promise(function(resolve) {
    res = s.replaceWithImage(new_url);
    if (res) {
      console.log('resolve');
      resolve();
    }
  })
}
async function createPdf() {
  return new Promise(function(resolve,reject) {
    let pdf =  DriveApp.getFileById(newSlideId).getBlob().getAs("application/pdf");
    pdf_name_full = pdf_file_name+data_name+'.pdf';
    pdf.setName(pdf_name_full);
    exportPdf = DriveApp.getFolderById(folderResponsePdfId).createFile(pdf);
    if (exportPdf) {
      resolve();
      console.log('Create PDF completed');
    } else {
      reject();
      console.log('Create PDF error');
    }
  });
}

async function sendEmailWithAttachment(email) {
  return new Promise(function(resolve,reject) {
    let file = DriveApp.getFolderById(folderResponsePdfId).getFilesByName(pdf_name_full);
    if (!file.hasNext()) {
      console.error("Could not open file "+pdf_name_full);
      return;
    }
    try {
      MailApp.sendEmail({
        to: email,
        subject: email_subject,
        htmlBody: email_message,
        attachments: [file.next().getAs(MimeType.PDF)]
      });
      resolve();
      console.log('sendEmailWithAttachment completed')
    } catch(e) {
      reject();
      console.log("sendEmailWithAttachment error with email (" + email + "). " + e);
    }
  });
}

function removeTempSlide() {
  try {
    DriveApp.getFileById(newSlideId).setTrashed(true);
    console.log('removeTempSlide completed');
  } catch(e) {
    console.log('removeTempSlide error')
  }
}

function updateStatusSent(numRow) {
  sheet.getRange(send_status_col.concat(numRow)).setValue(sent_status);
  console.log('updateStatusSent completed');
}

function formatUrlImg(url) {
  let new_url = '';
  let start_url = 'https://drive.google.com/uc?id=';
  new_url = start_url + getIdFromUrl(url);
  return new_url;
}

function getIdFromUrl(url) {
  return url.match(/[-\w]{25,}/);
}

function validateEmail(email) {
  var re = /\S+@\S+\.\S+/;
  if (!re.test(email)) {
    return false;
  } else {
    return true;
  }
}
