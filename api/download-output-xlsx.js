// Vercel Serverless Function: GET /api/download-output-xlsx
const fs = require('fs');
const path = require('path');

module.exports = (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  const filePath = path.join(process.cwd(), 'outputs', 'Doanh_Nghiep_Phan_Tich_Output.xlsx');
  if (fs.existsSync(filePath)) {
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="Doanh_Nghiep_Phan_Tich_Output.xlsx"');
    const stream = fs.createReadStream(filePath);
    return stream.pipe(res);
  } else {
    return res.status(404).send('File not found');
  }
};
