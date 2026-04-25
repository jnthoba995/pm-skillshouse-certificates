export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' })
  }

  return res.status(200).json({
    status: 'mock',
    message: 'Register review API ready. Document AI connection comes next.',
    rows: [
      {
        name: 'Sample',
        surname: 'Participant',
        idNumber: '',
        contact: '',
        email: '',
        gender: '',
        status: 'Review'
      }
    ]
  })
}
