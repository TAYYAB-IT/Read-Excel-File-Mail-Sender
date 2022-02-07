const express = require('express')
const app = express()
const excel=require('exceljs')
const mailer=require('nodemailer')
require('dotenv').config();
const sender=async(Name,Email)=>{
    const mail=  mailer.createTransport({
          service:'gmail',
        
          auth:{
              user:process.env.email,
              pass:process.env.password,
          }
      })
      
  
   const mail_details={
       from:process.env.email,
       to:Email,
       subject:"Testing Phase",
       //text:"Hey Buddy!"
       html:`<h1>Welcome! Mr. ${Name}</h1><a href='https://github.com/TAYYAB-IT/JS_Series'>Click Here</a>`,
   }
  
   
  await mail.sendMail(mail_details,(err, info)=>{
       if(err){console.log(err)}
       else{
           console.log("Email Sent, "+info.response)
       }
   })
  } 

  app.get('/email',async(req,res)=>{
const wb=new excel.Workbook();
wb.xlsx.readFile('./Files/Emails_List.xlsx').then(async()=>{
    const sh=wb.getWorksheet("Sheet1")
    for(var i=2;i<=sh.rowCount;i++){
        console.log(sh.getRow(i).getCell(2).value.text)
     await  sender(sh.getRow(i).getCell(1).value.toString(),sh.getRow(i).getCell(2).value.text).then(()=>{
          console.log("Sent");
       })
    }
    
})

  })
  app.listen(3000,()=>{
      console.log("Server is Active")
  })