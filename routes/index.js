var express = require('express');
var axios = require('axios');
var fs = require('fs');
var path = require('path')
var jwt = require('jsonwebtoken')
var formidable = require('formidable')
var router = express.Router()
var dayjs = require('dayjs')
var db = require("../db/db")

var PizZip = require('pizzip')
var Docxtemplater = require('docxtemplater')
var archiver=require("archiver")

var rootURL="";

var root = path.resolve(__dirname,'../')
var clone =(e)=> {
  return JSON.parse(JSON.stringify(e))
}

const SECRET_KEY = 'ANSAIR-SYSTEM'

var callSQLProc = (sql, params, res) => {
  return new Promise (resolve => {
    db.procedureSQL(sql,JSON.stringify(params),(err,ret)=>{
      if (err) {
        res.status(500).json({ code: -1, msg: '提交请求失败，请联系管理员！', data: null})
      }else{
        resolve(ret)
      }
    })
  })
}

var callP = async (sql, params, res) => {
  return  await callSQLProc(sql, params, res)
}


var decodeUser = (req)=>{
  let token = req.headers.authorization
  return  JSON.parse(token?.split(' ')[1])
}


router.post('/login',async (req, res, next) =>{
  
  let params = req.body
  let sql = `CALL PROC_LOGIN(?)`
  let r = await callP(sql, params, res)
  if (r.length > 0) {
    let ret = clone(r[0])
    let token = jwt.sign(ret, SECRET_KEY)
    res.status(200).json({code: 200, data: ret, token: token, msg: '登录成功'})
  } else {
    res.status(200).json({code: 301, data: null, msg: '用户名或密码错误'})
  }
})


router.post('/qryCls', async (req, res, next) =>{
  let uid = decodeUser(req).uid
  let params = {uid:uid}
  let sql= `CALL PROC_QRY_CLS(?)`
  let r = await callP(sql, params, res)
  res.status(200).json({ code: 200, data: r })
});

router.post('/qryClsSim', async (req, res, next) =>{
  let uid = decodeUser(req).uid

  let params = {uid:uid, code: req.body.code, term: req.body.term}
  let r = await callP(sql, params, res)
  res.status(200).json({ code: 200, data: r })
});

router.post('/qryClsMain', async (req, res, next) =>{
  let uid = decodeUser(req).uid
  let params = {uid:uid, code: req.body.code, term: req.body.term}

  let sql1= `CALL PROC_QRY_CLS_MAIN(?)`
  let sql2= `CALL PROC_QRY_TECH(?)`
  let sql3= `CALL PROC_QRY_EXP(?)`
  let sql6= `CALL PROC_QRY_CLS_BASE(?)`

  let r = await callP(sql1, params, res)
  let s = await callP(sql2, params, res)
  let t = await callP(sql3, params, res)

  res.status(200).json({ code: 200, data: r, tecList:s, expList:t })
});


router.post('/istClsHis', async (req, res, next) =>{
  let uid = decodeUser(req).uid

  req.body.uid = uid
  let params = req.body

  let sql1= `CALL PROC_INSERT_CLS(?)`
  let sql2= `CALL PROC_QRY_TECH(?)`
  let sql3= `CALL PROC_QRY_EXP(?)`
  let r = await callP(sql1, params, res)
  let s = await callP(sql2, params, res)
  let t = await callP(sql3, params, res)
  res.status(200).json({ code: 200, data: r, tecList:s, expList:t })
});

router.post('/savCls', async (req, res, next) =>{
  let uid = decodeUser(req).uid
  req.body.uid = uid
  let params = req.body

  let sql1= `CALL PROC_SAV_CLS(?)`
  let sql2= `CALL PROC_SAV_TECH(?)`
  let sql3= `CALL PROC_SAV_EXP(?)`
  let r = await callP(sql1, params, res)
  let s = await callP(sql2, params, res)
  let t = await callP(sql3, params, res)
  res.status(200).json({ code: 200, data: r, tecList:s, expList:t })
});

router.post('/qryClsHis', async (req, res, next) =>{
  let uid = decodeUser(req).uid

  req.body.uid = uid
  let params = req.body
  let sql= `CALL PROC_QRY_CLS_MAIN(?)`

  let r = await callP(sql, params, res)

  res.status(200).json({ code: 200, data: r})
});

router.post('/istClsSim', async (req, res, next) =>{
  let uid = decodeUser(req).uid

  req.body.uid = uid
  let params = req.body
  let sql1= `CALL PROC_INSERT_CLS_SIM(?)`
  let sql2= `CALL PROC_QRY_TECH(?)`
  let sql3= `CALL PROC_QRY_EXP(?)`
  let r = await callP(sql1, params, res)
  let s = await callP(sql2, params, res)
  let t = await callP(sql3, params, res)
  res.status(200).json({ code: 200, data: r, tecList:s, expList:t })
});



router.post('/qryClsBase', async (req, res, next) =>{
  let uid = decodeUser(req).uid
  let params = {uid:uid, code: req.body.code, term: req.body.term}
  let sql1= `CALL PROC_QRY_CLS_BASE(?)`
  let sql2= `CALL PROC_QRY_TECH(?)`
  let sql3= `CALL PROC_QRY_EXP(?)`
 

  let r = await callP(sql1, params, res)
  let s = await callP(sql2, params, res)
  let t = await callP(sql3, params, res)


    res.status(200).json({ code: 200, data: r, tecList:s, expList:t })

});

router.post('/qryGenDocx', async (req, res, next) =>{
  let sql  = `CALL PROC_GEN_DOC(?)`
  let r = await callP(sql, null, res)
  let product=[];
  let tech=[];
  let exp=[];
  for(let i in r) {  
    for(let j in r[i]){
      if(r[i][j]==''||r[i][j]===undefined||r[i][j]===null||r[i][j]==='undefined'||r[i][j]==='null')r[i][j]="无";
    }
  }
  let obj={};
  let i = 1;
  for(i=1;i<r.length;i++){

    obj={};
    obj['cls']=r[i-1].cls;
    obj['st_num']=r[i-1].st_num;
    obj['wt']=r[i-1].wt;
    obj['addr']=r[i-1].addr;

    if(r[i-1].name==r[i].name&&r[i-1].uid==r[i].uid){
      product.push(obj);
    }else{
      product.push(obj);
      let content = fs.readFileSync(path.resolve(__dirname, '../test.docx'), 'binary');
      let zip = new PizZip(content);
      let doc;
      try { doc = new Docxtemplater(zip) } catch (error) { errorHandler(error); }
      let params = {code:r[i-1].code, uid:r[i-1].uid};
      let sql1= `CALL PROC_QRY_EXP(?)`
      let sql2= `CALL PROC_QRY_TECH(?)`
      x = await callP(sql2, params, res);
      s = await callP(sql1, params, res);
    

      let curtime = dayjs().format('YYYY-MM-DD')

      if(x.length==0){
        let techobj={};
        techobj["教学周次"]="无";
        techobj["教学课时"]="无";
        techobj["教学内容"]="无";
        techobj["教学形式"]="无";
        techobj["作业辅导"]="无";
        tech.push(techobj);
      }
      
      if(s.length==0){
        let expobj={};
        expobj["实验周次"]="无";
        expobj["实验课时"]="无";
        expobj["实验项目"]="无";
        expobj["实验性质"]="无";
        expobj["实验要求"]="无";
        expobj["实验教室"]="无";
        expobj["每组人数"]="无";
        exp.push(expobj);
      }
      for(let j=0;j<x.length;j++){
        let techobj={};
        techobj["教学周次"]=`第${j+1}周`;
        techobj["教学课时"]=x[j].hour;
        techobj["教学内容"]=x[j].cnt;
        techobj["教学形式"]=x[j].method;
        techobj["作业辅导"]=x[j].task;
        tech.push(techobj);
      }

      for(let q=0;q<s.length;q++){
        let expobj={};
        expobj["实验周次"]=`第${q+1}周`;
        expobj["实验课时"]=s[q].hour;
        expobj["实验项目"]=s[q].name;
        expobj["实验性质"]=s[q].type;
        expobj["实验要求"]=s[q].prop;
        expobj["实验教室"]=s[q].addr;
        expobj["每组人数"]=s[q].gnum;
        exp.push(expobj);
      }
      doc.setData({
        课程名称:r[i-1].bname,
        学年时间:r[i-1].term,
        授课校区:r[i-1].pos,
        开课学院:r[i-1].col,
        课程学分:r[i-1].mark,
        教学周期:r[i-1].week,
        理论课时:r[i-1].t_hour,
        实验课时:r[i-1].e_hour,
        周学时数:r[i-1].t_hour+r[i-1].e_hour,
        总课时数:16*(r[i-1].t_hour+r[i-1].e_hour),
        主讲老师:r[i-1].m_tech,
        辅导教室:r[i-1].s_tech,
        答疑时间:r[i-1].q_time,
        答疑地点:r[i-1].q_addr,
        课程描述:r[i-1].desc,
        使用教材及参考书目:r[i-1].mate,
        课程考核:r[i-1].exam,
        教学方法与手段:r[i-1].method,
        签名:r[i-1].uname,
        product:product,
        tech:tech,
        exp:exp,
        时间:curtime,
      });

      try { doc.render() } catch (error) { errorHandler(error); }
      var buf = doc.getZip().generate({ type: 'nodebuffer' });
      r[i-1].bname = r[i-1].bname.replace(/\//g, " ");
      fs.writeFileSync(path.resolve(__dirname, `../build/docx/${r[i-1].bname}_${r[i-1].uname}.docx`), buf);
      product=[];
      exp=[];
      tech=[];
    }
    if(i==r.length-1){
      obj={};
      obj['cls']=r[i].cls;
      obj['st_num']=r[i].st_num;
      obj['wt']=r[i].wt;
      obj['addr']=r[i].addr;
    
      product.push(obj);
      let content = fs.readFileSync(path.resolve(__dirname, '../test.docx'), 'binary');
      let zip = new PizZip(content);
      let doc;
      try { doc = new Docxtemplater(zip) } catch (error) { errorHandler(error); }
      let params = {code:r[i].code, uid:r[i].uid};
      let sql1= `CALL PROC_QRY_EXP(?)`
      let sql2= `CALL PROC_QRY_TECH(?)`
      x = await callP(sql2, params, res);
      s = await callP(sql1, params, res);
    

      let curtime = dayjs().format('YYYY-MM-DD')

      if(x.length==0){
        let techobj={};
        techobj["教学周次"]="无";
        techobj["教学课时"]="无";
        techobj["教学内容"]="无";
        techobj["教学形式"]="无";
        techobj["作业辅导"]="无";
        tech.push(techobj);
      }
      if(s.length==0){
        let expobj={};
        expobj["实验周次"]="无";
        expobj["实验课时"]="无";
        expobj["实验项目"]="无";
        expobj["实验性质"]="无";
        expobj["实验要求"]="无";
        expobj["实验教室"]="无";
        expobj["每组人数"]="无";
        exp.push(expobj);
      }

      for(let j=0;j<x.length;j++){
        let techobj={};
        techobj["教学周次"]=`第${j+1}周`;
        techobj["教学课时"]=x[j].hour;
        techobj["教学内容"]=x[j].cnt;
        techobj["教学形式"]=x[j].method;
        techobj["作业辅导"]=x[j].task;
        tech.push(techobj);
      }

      for(let q=0;q<s.length;q++){
        let expobj={};
        expobj["实验周次"]=`第${q+1}周`;
        expobj["实验课时"]=s[q].hour;
        expobj["实验项目"]=s[q].name;
        expobj["实验性质"]=s[q].type;
        expobj["实验要求"]=s[q].prop;
        expobj["实验教室"]=s[q].addr;
        expobj["每组人数"]=s[q].gnum;
        exp.push(expobj);
      }
      doc.setData({
        课程名称:r[i].bname,
        学年时间:r[i].term,
        授课校区:r[i].pos,
        开课学院:r[i].col,
        课程学分:r[i].mark,
        教学周期:r[i].week,
        理论课时:r[i].t_hour,
        实验课时:r[i].e_hour,
        周学时数:r[i].t_hour+r[i].e_hour,
        总课时数:16*(r[i].t_hour+r[i].e_hour),
        主讲老师:r[i].m_tech,
        辅导教室:r[i].s_tech,
        答疑时间:r[i].q_time,
        答疑地点:r[i].q_addr,
        课程描述:r[i].desc,
        使用教材及参考书目:r[i].mate,
        课程考核:r[i].exam,
        教学方法与手段:r[i].method,
        签名:r[i].uname,
        product:product,
        tech:tech,
        exp:exp,
        时间:curtime,
      });

      try { doc.render() } catch (error) { errorHandler(error); }
      var buf = doc.getZip().generate({ type: 'nodebuffer' });

      fs.writeFileSync(path.resolve(__dirname, `../build/docx/${r[i].bname}_${r[i].uname}.docx`), buf);
      product=[];
      exp=[];
      tech=[];
    }
  }

    

  const dir = path.resolve(__dirname, '../build/docx/')
  const files = fs.readdirSync(dir)
  const output = fs.createWriteStream(path.join(__dirname, '../build/教学安排表.zip'), { encoding: 'utf8' });
  const archive = archiver('zip', { zlib: { level: 9 } }); 

  archive.pipe(output); 
  for(let i=0; i<files.length; i++){
    var stream = fs.createReadStream(path.resolve(__dirname, '../build/docx/', files[i])); 
    archive.append(stream, { name: files[i] }); 
  }
  archive.finalize(); 

  res.status(200).json({ code: 200, path: path.resolve(__dirname + "/教学安排表.zip") });
  
  
});




router.post('/upload', function (req, res, next) {
  const form = formidable({uploadDir: `${__dirname}/../img`});
  form.on('fileBegin', function (name, file) {
    if (req.body.uid) {
      file.filepath = `img/${req.body.uid}.jpeg`;
    }
  });
 
  form.parse(req, (err, fields, files) => {
    if (err) {
      next(err);
      return;
    }
    res.status(200).json({
      code: 200,
      msg: '上传照片成功',
      data: {path: files.file.filepath}
    })
  });
});


module.exports = router