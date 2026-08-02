/**
 * Places.gs
 * HotelDatabase Google Places API integration
 */

const PLACES_CONFIG = {
  API_KEY: PropertiesService.getScriptProperties().getProperty('GOOGLE_PLACES_API_KEY'),
  LANGUAGE: 'ja',
  REGION: 'jp'
};

function searchPlaceFromGoogle(query) {
  const url = 'https://maps.googleapis.com/maps/api/place/textsearch/json';
  const params = {
    query: query,
    language: PLACES_CONFIG.LANGUAGE,
    region: PLACES_CONFIG.REGION,
    key: PLACES_CONFIG.API_KEY
  };
  const res = UrlFetchApp.fetch(url + '?' + buildQuery(params));
  const json = JSON.parse(res.getContentText());
  return json.results || [];
}

function getPlaceDetails(placeId) {
  const url = 'https://maps.googleapis.com/maps/api/place/details/json';
  const params = {
    place_id: placeId,
    language: PLACES_CONFIG.LANGUAGE,
    fields: 'name,formatted_address,formatted_phone_number,website,rating,user_ratings_total,geometry',
    key: PLACES_CONFIG.API_KEY
  };
  const res = UrlFetchApp.fetch(url + '?' + buildQuery(params));
  const json = JSON.parse(res.getContentText());
  return json.result || null;
}

function updateFacilityRow(sheet,rowNumber,placeData){
  sheet.getRange(rowNumber,1,1,6).setValues([[
    placeData.name || '',
    placeData.formatted_address || '',
    placeData.formatted_phone_number || '',
    placeData.website || '',
    placeData.rating || '',
    placeData.user_ratings_total || ''
  ]]);
}

function normalizePlaceData(place){
  return {
    name: place.name || '',
    address: place.formatted_address || '',
    phone: place.formatted_phone_number || '',
    website: place.website || '',
    rating: place.rating || 0,
    reviews: place.user_ratings_total || 0
  };
}

function buildQuery(params){
  return Object.keys(params).map(function(k){
    return encodeURIComponent(k)+'='+encodeURIComponent(params[k]);
  }).join('&');
}

function testPlacesAPI(){
  Logger.log(searchPlaceFromGoogle('東京 ホテル'));
}
