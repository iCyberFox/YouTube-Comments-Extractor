exports.handler = async function(event, context) {
  return {
    statusCode: 200,
    body: JSON.stringify({
      apiKey: process.env.VITE_YOUTUBE_API_KEY
    })
  };
};