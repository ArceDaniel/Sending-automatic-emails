const { Configuration, OpenAIApi } = require("openai");
const nodemailer = require("nodemailer");
const xlsx = require("xlsx");

// configuramos nodemailer
const transporter = nodemailer.createTransport({
  host: "smtp.gmail.com",
  port: 465,
  secure: true,
  auth: {
    user: "tu email",
    pass: "tu pass de google, esta se obtiene yendo a la configuracion de tu cuenta y busar tu clave  https://myaccount.google.com/u/1/apppasswords?rapt=AEjHL4MkYsBQIIofXxGGfeTTXtLX1C2fhg6vB48siZadGiipM4rRbSLi303ynu85JvnuZuV5Rv0soMZcFjpTsKk47lRa7Vbk-g",
  },
});


//configuramos chat gpt
const configuration = new Configuration({
  apiKey: "key de chatGPT https://platform.openai.com/account/api-keys",
});
const openai = new OpenAIApi(configuration);

//Mapeamos el excel y lo devolvemos en formato array

async function dataOfExel() {
  //aca buscamos el excel
  const workbook = xlsx.readFile("postulaciones.xlsx");
  //aca especifacamos que hoja del excel mapear
  const sheet = workbook.Sheets["HenryProjects"];
  const data = xlsx.utils.sheet_to_json(sheet);
  // console.log(data[0]);
  return data;
}

async function sendMail() {
  //obtenemos el array
  const excel = await dataOfExel();

  // un bucle for para enviar mail por mail
  for (let i = 0; i < excel.length; i++) {
    const data = excel[i];

    const completion = await openai.createChatCompletion({
      model: "gpt-3.5-turbo",
      messages: [
        {
          role: "user",
          ///aca personalicen su promp, les aconsejo que sean mas detallados, yo lo hice asi nomas
          content: `Hola mi nombre es Jonathan Daniel Arce, necesito enviar un correo electrónico a para postularme como fullStack / Backend a ${data["Nombre"]}, que es un ${data["Tipo de institución"]} que se dedica a ${data["Industria / Sector"]} donde es su postulacion tiene la siguiente descripcion ${data["Descripcion Proyecto"]}, datos adicionales sobre mi, tengo grandes conocimentos en el backend, me especializo en desarrollo de api solidas y segura, utilizo typescript y tengo conocimientos sobre bases de datos sql y no sql`,
        },
      ],
    });
    const text = completion.data.choices[0].message.content;


    //envimos los mails
    const mailOptions = {
      from: "tu email",
      to: data["Link o mail de aplicacion"],
      subject: `${data["Nombre"]} Arce Daniel-Postulación Henry project.`,
      text,
      attachments: [
        {
          // aca buscamos nuestro cv, puede estar en formato pdf
          filename: "cv.docx",
          path: "./cv.docx",
        },
      ],
    };

    transporter.sendMail(mailOptions, function (error, info) {
      if (error) {
        console.log(error);
      } else {
        console.log("Correo electrónico enviado: " + info.response);
      }
    });
  }
}


  sendMail();
