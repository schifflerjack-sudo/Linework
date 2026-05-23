const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxOifhMJ6033zmaMRsC45Q7byW8jfoxKwNYkWgyw51S-ZWPQUbSZchfHdrdVGsmC_eMFw/exec';

export async function onRequestPost(context) {
  try {
    const formData = await context.request.formData();
    const params = new URLSearchParams();
    for (const [key, value] of formData.entries()) {
      params.append(key, value);
    }

    await fetch(SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString(),
      redirect: 'follow'
    });

    return new Response(JSON.stringify({ status: 'success' }), {
      headers: { 'Content-Type': 'application/json' }
    });
  } catch (err) {
    return new Response(JSON.stringify({ status: 'error', message: err.message }), {
      status: 500,
      headers: { 'Content-Type': 'application/json' }
    });
  }
}
