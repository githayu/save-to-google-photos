declare const OAuth2: any

/**
 * Authorizes and makes a request to the Google Photos API.
 */
const auth = () => {
  const service = getService()
  const authorizationUrl = service.getAuthorizationUrl()

  console.error('認証URL:', authorizationUrl)
}

function doPost(e: any) {
  let status = false

  console.log('PostData:', e.parameter)

  try {
    saveToPhotos(e.parameter)

    status = true
  } catch (err) {
    console.error(err)
  }

  return ContentService.createTextOutput(
    JSON.stringify({
      status,
    })
  )
}

function saveToPhotos({
  url,
  name,
  description,
  albumName,
}: {
  url: string
  name?: string
  description?: string
  albumName?: string
}) {
  const service = getService()

  if (service.hasAccess()) {
    if (!url) {
      throw new Error('URL is invalid')
    }

    const accessToken = 'Bearer ' + service.getAccessToken()
    const payload = fetchImage(url)
    const uploadToken = uploadBytes({
      accessToken,
      name,
      payload,
    })

    let album

    if (albumName) {
      ;[album] = getAlbums(accessToken).albums.filter(
        (album: { title: string }) => album.title === albumName
      )

      if (!album) {
        album = createAlbum(accessToken, albumName)
      }
    }

    console.log('Album:', album)

    const response = createMediaItem({
      accessToken,
      uploadToken,
      description,
      albumId: album.id,
    })

    console.log('MediaItem:', response)
  } else {
    auth()
  }
}

function getAlbums(accessToken: string) {
  const res = UrlFetchApp.fetch(
    'https://photoslibrary.googleapis.com/v1/albums',
    {
      headers: {
        Authorization: accessToken,
      },
    }
  )

  const statusCode = res.getResponseCode()

  if (statusCode === 200) {
    return JSON.parse(res.getContentText())
  } else {
    throw new Error(`GetAlbums: ${statusCode}`)
  }
}

function createAlbum(accessToken: string, title: string) {
  const res = UrlFetchApp.fetch(
    'https://photoslibrary.googleapis.com/v1/albums',
    {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Authorization: accessToken,
      },
      payload: JSON.stringify({
        album: { title },
      }),
    }
  )

  const statusCode = res.getResponseCode()

  if (statusCode === 200) {
    return JSON.parse(res.getContentText())
  } else {
    throw new Error(`CreateAlbum: ${statusCode}`)
  }
}

function uploadBytes({
  accessToken,
  name,
  payload,
}: {
  accessToken: string
  name?: string
  payload: GoogleAppsScript.Base.Blob
}) {
  const headers: {
    [x: string]: string
  } = {
    Authorization: accessToken,
    'X-Goog-Upload-Protocol': 'raw',
  }

  if (name) {
    headers['X-Goog-Upload-File-Name'] = name
  }

  const res = UrlFetchApp.fetch(
    'https://photoslibrary.googleapis.com/v1/uploads',
    {
      method: 'post',
      contentType: 'application/octet-stream',
      headers,
      payload,
    }
  )

  const statusCode = res.getResponseCode()

  if (statusCode === 200) {
    return res.getContentText()
  } else {
    throw new Error(`UploadBytes: ${statusCode}`)
  }
}

function createMediaItem({
  accessToken,
  uploadToken,
  description,
  albumId,
}: {
  accessToken: string
  uploadToken: string
  description?: string
  albumId?: string
}) {
  const res = UrlFetchApp.fetch(
    'https://photoslibrary.googleapis.com/v1/mediaItems:batchCreate',
    {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Authorization: accessToken,
      },
      payload: JSON.stringify({
        albumId,
        newMediaItems: [
          {
            description,
            simpleMediaItem: {
              uploadToken,
            },
          },
        ],
      }),
    }
  )

  const statusCode = res.getResponseCode()

  if (statusCode === 200) {
    return JSON.parse(res.getContentText())
  } else {
    throw new Error(`CreateMediaItem: ${statusCode}`)
  }
}

function fetchImage(url: string) {
  const res = UrlFetchApp.fetch(url)
  const statusCode = res.getResponseCode()

  if (statusCode === 200) {
    return res.getBlob()
  } else {
    throw new Error(`FetchImage: ${statusCode}`)
  }
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  getService().reset()
}

/**
 * Configures the service.
 */
function getService() {
  const scriptProps = PropertiesService.getScriptProperties()

  return (
    OAuth2.createService('Photos')
      // Set the endpoint URLs.
      .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/v2/auth')
      .setTokenUrl('https://oauth2.googleapis.com/token')

      // Set the client ID and secret.
      .setClientId(scriptProps.getProperty('CLIENT_ID'))
      .setClientSecret(scriptProps.getProperty('CLIENT_SECRET'))

      // Set the name of the callback function that should be invoked to
      // complete the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(scriptProps)

      // Set the scope and additional Google-specific parameters.
      .setScope('https://www.googleapis.com/auth/photoslibrary')
      .setParam('access_type', 'offline')
      .setParam('prompt', 'consent')
      .setParam('login_hint', scriptProps.getProperty('EMAIL'))
  )
}

/**
 * Handles the OAuth callback.
 */
function authCallback(request: any) {
  const service = getService()
  const authorized = service.handleCallback(request)

  if (authorized) {
    return HtmlService.createHtmlOutput('Success!')
  } else {
    return HtmlService.createHtmlOutput('Denied.')
  }
}

/**
 * Logs the redict URI to register in the Google Developers Console.
 */
function logRedirectUri() {
  console.log(OAuth2.getRedirectUri())
}
