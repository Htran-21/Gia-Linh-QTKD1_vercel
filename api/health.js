module.exports = (req, res) => {
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }
  return res.status(200).json({
    status: 'ok',
    system: 'Gia Linh Data Studio',
    version: '2.5',
    deployed_on: 'Vercel Serverless'
  });
};
