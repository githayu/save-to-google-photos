declare const OAuth2: any

/*
 * This sample demonstrates how to configure the library for Google APIs, using
 * end-user authorization (Web Server flow).
 * https://developers.google.com/identity/protocols/OAuth2WebServer
 */
function test() {
  doGet({
    parameter: {
      url: 'https://via.placeholder.com/150',
      name: 'name',
      description: 'description',
      albumName: 'albumName',
    },
  })
}

/**
 * Authorizes and makes a request to the Google Drive API.
 */
function doGet(e: any) {
  let status = false

  try {
    saveToPhotos(e.parameter)

    status = true
  } catch (err) {
    console.error(err)
  }

  return ContentService.createTextOutput(
    JSON.stringify({
      status: true,
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

  if (service.hasAccess() && url) {
    const accessToken = 'Bearer ' + service.getAccessToken()
    const payload = fetchImage(url)
    const uploadToken = uploadBytes({
      accessToken,
      name,
      payload,
    })

    console.log('uploadToken', uploadToken)

    let album

    if (albumName) {
      ;[album] = getAlbums(accessToken).albums.filter(
        (album: { title: string }) => album.title === albumName
      )

      if (!album) {
        album = createAlbum(accessToken, albumName)
      }
    }

    const response = createMediaItem({
      accessToken,
      uploadToken,
      description,
      albumId: album ? album.id : undefined,
    })

    console.log('createMediaItem', response)
  } else {
    const authorizationUrl = service.getAuthorizationUrl()

    console.warn(
      `Open the following URL and re-run the script: ${authorizationUrl}`
    )
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
    throw new Error(`GetAlbumError: ${statusCode}`)
  }
}

function createAlbum(accessToken: string, title: string) {
  const res = UrlFetchApp.fetch(
    'https://photoslibrary.googleapis.com/v1/albums',
    {
      method: 'post',
      headers: {
        Authorization: accessToken,
        'Content-Type': 'application/json',
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
    throw new Error(`CreateAlbumError: ${statusCode}`)
  }
}

function uploadBytes({
  accessToken,
  name,
  payload,
}: {
  accessToken: string
  name?: string
  payload: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions['payload']
}) {
  const headers: {
    [x: string]: string
  } = {
    Authorization: accessToken,
    'Content-Type': 'application/octet-stream',
    'X-Goog-Upload-Protocol': 'raw',
  }

  if (name) {
    headers['X-Goog-Upload-File-Name'] = name
  }

  const res = UrlFetchApp.fetch(
    'https://photoslibrary.googleapis.com/v1/uploads',
    {
      method: 'post',
      headers,
      payload,
    }
  )

  const statusCode = res.getResponseCode()

  if (statusCode === 200) {
    return res.getContentText()
  } else {
    throw new Error(`UploadBytesError: ${statusCode}`)
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
      headers: {
        Authorization: accessToken,
        'Content-Type': 'application/json',
      },
      payload: JSON.stringify({
        albumId,
        newMediaItems: [
          {
            description: description,
            simpleMediaItem: {
              uploadToken: uploadToken,
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
    throw new Error(`createMediaItemError: ${statusCode}`)
  }
}

function fetchImage(url: string) {
  const res = UrlFetchApp.fetch(url)
  const statusCode = res.getResponseCode()

  if (statusCode === 200) {
    return res.getBlob()
  } else {
    throw new Error(`fetchImageError: ${statusCode}`)
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
    OAuth2.createService('Google')
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
      .setPropertyStore(PropertiesService.getScriptProperties())

      // Set the scope and additional Google-specific parameters.
      .setScope('https://www.googleapis.com/auth/photoslibrary')
      .setParam('access_type', 'offline')
      .setParam('approval_prompt', 'auto')
      .setParam('login_hint', Session.getActiveUser().getEmail())
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
