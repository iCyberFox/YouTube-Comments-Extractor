exports.handler = async function (event) {
  const videoId = event.queryStringParameters?.videoId;

  if (!videoId) {
    return {
      statusCode: 400,
      body: JSON.stringify({ error: 'Missing videoId' })
    };
  }

  try {
    const watchPageResponse = await fetch(
      `https://www.youtube.com/watch?v=${encodeURIComponent(videoId)}&bpctr=9999999999&hl=uk&gl=UA`,
      {
        headers: {
          'User-Agent': 'Mozilla/5.0'
        }
      }
    );

    if (!watchPageResponse.ok) {
      throw new Error(`Watch page request failed: ${watchPageResponse.status}`);
    }

    const watchPageHtml = await watchPageResponse.text();
    const apiKeyMatch = watchPageHtml.match(/"INNERTUBE_API_KEY":"([^"]+)"/);
    const clientVersionMatch = watchPageHtml.match(/"INNERTUBE_CLIENT_VERSION":"([^"]+)"/);

    if (!apiKeyMatch || !clientVersionMatch) {
      throw new Error('Could not resolve YouTube internal client data');
    }

    const apiKey = apiKeyMatch[1];
    const clientVersion = clientVersionMatch[1];

    const comments = [];
    const seenCommentIds = new Set();
    let continuation = null;
    let requests = 0;

    while (requests < 20) {
      const payload = {
        context: {
          client: {
            clientName: 'WEB',
            clientVersion,
            hl: 'uk',
            gl: 'UA'
          }
        },
        videoId,
        params: 'CgZjb21tZW50cw%3D%3D'
      };

      if (continuation) {
        payload.continuation = continuation;
      }

      const response = await fetch(
        `https://www.youtube.com/youtubei/v1/next?prettyPrint=false&key=${apiKey}`,
        {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'User-Agent': 'Mozilla/5.0'
          },
          body: JSON.stringify(payload)
        }
      );

      if (!response.ok) {
        throw new Error(`Internal comments request failed: ${response.status}`);
      }

      const data = await response.json();
      const mutations = data?.frameworkUpdates?.entityBatchUpdate?.mutations || [];

      for (const mutation of mutations) {
        const payloadItem = mutation.payload?.commentEntityPayload;
        if (!payloadItem) {
          continue;
        }

        const properties = payloadItem.properties || {};
        const replyLevel = Number(properties.replyLevel || 0);
        const commentId = properties.commentId || payloadItem.key;
        const content = properties.content?.content || '';
        const author = payloadItem.author || {};
        const displayName = author.displayName || '';
        const publishedTime = properties.publishedTime || '';

        if (!content || !displayName || !commentId || replyLevel > 0) {
          continue;
        }

        if (seenCommentIds.has(commentId)) {
          continue;
        }

        seenCommentIds.add(commentId);
        comments.push({
          commentId,
          name: displayName,
          text: content,
          date: publishedTime,
          replyLevel
        });
      }

      const nextContinuation = findContinuationToken(data);
      if (!nextContinuation) {
        break;
      }

      continuation = nextContinuation;
      requests += 1;
    }

    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'application/json; charset=utf-8'
      },
      body: JSON.stringify({
        comments,
        total: comments.length
      })
    };
  } catch (error) {
    return {
      statusCode: 500,
      body: JSON.stringify({
        error: error.message
      })
    };
  }
};

function findContinuationToken(value) {
  if (!value || typeof value !== 'object') {
    return null;
  }

  if (Array.isArray(value)) {
    for (const item of value) {
      const found = findContinuationToken(item);
      if (found) {
        return found;
      }
    }
    return null;
  }

  if (value.continuationCommand && typeof value.continuationCommand === 'object') {
    const token = value.continuationCommand.token;
    const requestType = value.continuationCommand.request;
    if (typeof token === 'string' && token && (!requestType || requestType === 'CONTINUATION_REQUEST_TYPE_WATCH_NEXT')) {
      return token;
    }
  }

  for (const key of Object.keys(value)) {
    const found = findContinuationToken(value[key]);
    if (found) {
      return found;
    }
  }

  return null;
}
