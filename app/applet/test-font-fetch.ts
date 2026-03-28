import fetch from 'node-fetch';
async function run() {
  try {
    const res1 = await fetch('https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Regular.ttf');
    console.log('Regular:', res1.status, res1.headers.get('content-type'));
    const res2 = await fetch('https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Medium.ttf');
    console.log('Medium:', res2.status, res2.headers.get('content-type'));
  } catch (e) {
    console.error(e);
  }
}
run();
