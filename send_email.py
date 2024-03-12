import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
import dotenv
import logging


logging.basicConfig(filename='email.log', level=logging.DEBUG,
                    format='%(asctime)s %(levelname)s %(name)s %(message)s')

# Define a logger for your application
logger = logging.getLogger(__name__)

dotenv.load_dotenv('settings.env')

def send_mail(patient_id, patient_name, test, clin_name, org_name, samp_col_date, samp_recd_date,TAT_date, email, bcc):
  sender_email = os.environ.get("SENDER_EMAIL")

  message = MIMEMultipart("alternative")
  message["From"] = sender_email
  message["To"] = email 
  message["Subject"] = f"Sample for {test} received at Greenarray"
 
  html = f"""\
<!doctype html>
<body style="word-spacing:normal;">
  <style type="text/css">
    outlook a {{
      padding: 0;
    }}

    body {{
      margin: 0;
      padding: 0;
      -webkit-text-size-adjust: 100%;
      -ms-text-size-adjust: 100%;
    }}

    table,
    td {{
      border-collapse: collapse;
      mso-table-lspace: 0pt;
      mso-table-rspace: 0pt;
    }}

    img {{
      border: 0;
      height: auto;
      line-height: 100%;
      outline: none;
      text-decoration: none;
      -ms-interpolation-mode: bicubic;
    }}

    p {{
      display: block;
      margin: 13px 0;
    }}
  </style>
  <style type="text/css">
    @media only screen and (min-width:480px) {{
      .mj-column-per-100 {{
        width: 100% !important;
        /* max-width: 100%; */
      }}
    }}
  </style>
  <style media="screen and (min-width:480px)">
    .moz-text-html .mj-column-per-100 {{
      width: 100% !important;
      /* max-width: 100%; */
    }}
  </style>
  <style type="text/css">
    @media only screen and (max-width:480px) {{
      table.mj-full-width-mobile {{
        width: 100% !important;
      }}

      td.mj-full-width-mobile {{
        width: auto !important;
      }}
    }}
  </style>

  <div style="">
    <!--[if mso | IE]><table align="center" border="0" cellpadding="0" cellspacing="0" class="" style="width:600px;" width="600" ><tr><td style="line-height:0px;font-size:0px;mso-line-height-rule:exactly;"><![endif]-->
    <div style="margin:0px auto">
      <table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="width:100%;">
        <tbody>
          <tr>
            <td style="direction:ltr;font-size:0px;padding:20px 0;text-align:center;">
              <!--[if mso | IE]><table role="presentation" border="0" cellpadding="0" cellspacing="0"><tr><td class="" style="vertical-align:top;width:600px;" ><![endif]-->
              <div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;">
                <table border="0" cellpadding="0" cellspacing="0" role="presentation" style="vertical-align:top;" width="100%">
                  <tbody>
                    <tr>
                      <td align="left" style="font-size:0px;padding:10px 25px;word-break:break-word;">
                        <div style="font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:16px;font-weight:900;line-height:1;text-align:left;color:#000000;">Dear {clin_name},</div>
                      </td>
                    </tr>
                    <tr>
                      <td align="left" style="font-size:0px;padding:10px 25px;padding-bottom:0px;word-break:break-word;">
                        <div style="font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:16px;line-height:1;text-align:left;color:#000000;">Your patient's sample has been registered for {test}. See the details below:</div>
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <!--[if mso | IE]></td></tr></table><![endif]-->
            </td>
          </tr>
        </tbody>
      </table>
    </div>
    <!--[if mso | IE]></td></tr></table><table align="center" border="0" cellpadding="0" cellspacing="0" class="" style="width:600px;" width="600" ><tr><td style="line-height:0px;font-size:0px;mso-line-height-rule:exactly;"><![endif]-->
    <div style="margin:0px auto;">
      <table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="width:100%;">
        <tbody>
          <tr>
            <td style="direction:ltr;font-size:0px;padding:20px 0;text-align:center;">
              <!--[if mso | IE]><table role="presentation" border="0" cellpadding="0" cellspacing="0"><tr><td class="" style="vertical-align:top;width:600px;" ><![endif]-->
              <div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;">
                <table border="0" cellpadding="0" cellspacing="0" role="presentation" style="vertical-align:top;">
                  <tbody>
                    <tr>
                      <td align="left" style="font-size:0px;padding:10px 25px;padding-bottom:0;word-break:break-word;">
                        <table cellpadding="0" cellspacing="0" width="100%"  style="color:#000000;font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:16px;line-height:22px;table-layout:auto;width:100%;border:none;">
                          <tr style="padding-bottom:100px">
                            <th style="text-align: left;">Patient Name: </th>
                            <td style="text-align: left">{patient_name}</td>
                          </tr>
                          <tr style="padding-bottom:100px">
                            <th style="text-align: left;">Clinician Name: </th>
                            <td style="text-align: left;">{clin_name}</td>
                          </tr>
                          <tr style="padding-bottom:10px">
                            <th style="text-align: left;">Organisation Name: </th>
                            <td style="text-align: left;">{org_name}</td>
                          </tr>
                          <tr style="padding-bottom:10px">
                            <th style="text-align: left;">Sample Collection Date: </th>
                            <td style="text-align: left;">{str(samp_col_date).split(' ')[0]}</td>
                          </tr>
                          <tr style="padding-bottom:10px">
                            <th style="text-align: left;">Sample Received Date: </th>
                            <td style="text-align: left;">{str(samp_recd_date).split(' ')[0]}</td>
                          </tr>
                          <tr style="padding-bottom:10px">
                            <th style="text-align: left; padding-right:20px">Expected Reporting Date:  </th>
                            <td style="text-align: left;">{str(TAT_date).split(' ')[0]}</td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <!--[if mso | IE]></td></tr></table><![endif]-->
            </td>
          </tr>
        </tbody>
      </table>
    </div>
    <!--[if mso | IE]></td></tr></table><table align="center" border="0" cellpadding="0" cellspacing="0" class="" style="width:600px;" width="600" ><tr><td style="line-height:0px;font-size:0px;mso-line-height-rule:exactly;"><![endif]-->
    <div style="margin:0px auto;">
      <table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="width:100%;">
        <tbody>
          <tr>
            <td style="direction:ltr;font-size:0px;padding:20px 0;text-align:center;">
              <!--[if mso | IE]><table role="presentation" border="0" cellpadding="0" cellspacing="0"><tr><td class="" style="vertical-align:top;width:600px;" ><![endif]-->
              <div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;">
                <table border="0" cellpadding="0" cellspacing="0" role="presentation" style="vertical-align:top;" width="100%">
                  <tbody>
                    <tr>
                      <td align="left" style="font-size:0px;padding:10px 25px;word-break:break-word;">
                        <div style="font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:14px;line-height:1;text-align:left;color:#000000;">While every effort will be made to ensure that the report is sent within TAT, in the case of unexpected delay, an intimation will be sent prior to the delivery date.
                          <mj-spacer height="50px" />
                        </div>
                      </td>
                    </tr>
                    <tr>
                      <td align="left" style="font-size:0px;padding:10px 25px;word-break:break-word;">
                        <div style="font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:14px;line-height:1;text-align:left;color:#000000;">1. Pretest genetic counseling is highly recommended. 2. In rare cases, the reporting may be delayed due to unforeseen circumstances. In such cases, you will receive an update. 3. In some cases, quality checks may fail due to sample type, DNA quality, or other reasons. In such cases, samples will not be processed further, and you will be updated.</div>
                      </td>
                    </tr>
                    <tr>
                      <td align="left" style="font-size:0px;padding:10px 25px;word-break:break-word;">
                        <div style="font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:14px;line-height:1;text-align:left;color:#000000;">Thanks & Regards</div>
                      </td>
                    </tr>
                    <tr>
                      <td align="left" style="font-size:0px;padding:10px 25px;padding-top:0px;word-break:break-word;">
                        <div style="font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:14px;line-height:1;text-align:left;color:#000000;">GreenArray Genomic Research & Solutions PVT LTD, Pune</div>
                      </td>
                    </tr>
                    <tr>
                      <td align="left" style="font-size:0px;padding:10px 25px;word-break:break-word;">
                        <div style="font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:13px;line-height:1;text-align:left;color:#000000;">--</div>
                      </td>
                    </tr>
                    <tr>
                      <td align="left" style="font-size:0px;padding:10px 25px;word-break:break-word;">
                        <div style="font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:14px;line-height:1;text-align:left;color:gray;">Thank you and regards,</div>
                      </td>
                    </tr>
                    <tr>
                      <td align="left" style="font-size:0px;padding:10px 25px;padding-top:0px;word-break:break-word;">
                        <div style="font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:14px;line-height:1;text-align:left;color:gray;">Team Greenarray</div>
                      </td>
                    </tr>
                    <tr>
                      <td align="left" style="font-size:0px;padding:10px 25px;word-break:break-word;">
                        <table border="0" cellpadding="0" cellspacing="0" role="presentation" style="border-collapse:collapse;border-spacing:0px;">
                          <tbody>
                            <tr>
                              <td style="width:148px;">
                                <img height="66" src="https://www.greenarraylabs.com/images/greenarray-logo.jpg" style="border:0;display:block;outline:none;text-decoration:none;height:66px;width:100%;font-size:13px;" width="148" />
                              </td>
                            </tr>
                          </tbody>
                        </table>
                      </td>
                    </tr>
                    <tr>
                      <td align="left" style="font-size:0px;padding:10px 25px;word-break:break-word;">
                        <div style="font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:13px;line-height:1;text-align:left;color:#000000;">Greenarray Genomic Research & Solutions of ADPL,</div>
                      </td>
                    </tr>
                    <tr>
                      <td align="left" style="font-size:0px;padding:10px 25px;padding-top:2px;word-break:break-word;">
                        <div style="font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:13px;line-height:1;text-align:left;color:#000000;">Office number- 201-206,</div>
                      </td>
                    </tr>
                    <tr>
                      <td align="left" style="font-size:0px;padding:10px 25px;padding-top:2px;word-break:break-word;">
                        <div style="font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:13px;line-height:1;text-align:left;color:#000000;">2ndFloor, Above P. N. Gadgil Showroom,</div>
                      </td>
                    </tr>
                    <tr>
                      <td align="left" style="font-size:0px;padding:10px 25px;padding-top:2px;word-break:break-word;">
                        <div style="font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:13px;line-height:1;text-align:left;color:#000000;">Happy Colony, Kothrud, Pune-411038.</div>
                      </td>
                    </tr>
                    <tr>
                      <td align="left" style="font-size:0px;padding:10px 25px;padding-top:2px;word-break:break-word;">
                        <div style="font-family:Ubuntu, Helvetica, Arial, sans-serif;font-size:13px;line-height:1;text-align:left;color:#000000;">Maharashtra, India.9823049121</div>
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <!--[if mso | IE]></td></tr></table><![endif]-->
            </td>
          </tr>
        </tbody>
      </table>
    </div>
    <!--[if mso | IE]></td></tr></table><![endif]-->
  </div>
</body>

</html>
  """

  part2 = MIMEText(html, "html")
  message.attach(part2)

  port = os.environ.get('PORT')
  password = os.environ.get('PASS')
  context = ssl.create_default_context()

  with smtplib.SMTP_SSL("smtp.gmail.com", port, context = context) as server:
      server.login(sender_email, password)
      try:
        server.sendmail(sender_email, [email] + bcc, message.as_string())
        logger.info(msg=f'Email ({email}) sent successfully')
      except (smtplib.SMTPException, smtplib.SMTPSenderRefused, smtplib.SMTPAuthenticationError) as e:
        logger.error(msg=f'Error sending email to {email}: {e}')
        
      