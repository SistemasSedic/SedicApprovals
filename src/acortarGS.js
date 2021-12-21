/**
 * Permite acortar las url
 * @param {string} url url
 * @returns URL acortada con tinyURL
 */
function Acortar(url) {
    url = 'http://tinyurl.com/api-create.php?url=' + url;
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    return response.getContentText()
}