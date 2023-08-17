/**
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/
/**
 * binary array to int
 */
function binArraytoInt(array) {
   return array.reduce((acc, val) => {
      return (acc << 1) | val;
   });
}

/**
 * 'yyyy-mm-dd' date String
 */
function _getNowDateISOFormattedString(){
  return _getISOTimeZoneCorrectedDateString(new Date());
}

/**
 * javascript toISOString timezone treatment
 */
function _getISOTimeZoneCorrectedDateString(date) {
  // timezone offset 처리 
  var tzoffset = date.getTimezoneOffset() * 60000; //offset in milliseconds
  return (new Date(date.getTime() - tzoffset)).toISOString().slice(0, -1).substring(0, 10);
}

/**
 * Simple string hash for checking two string difference
 */
function hash(str) {
  var hash = 0,
  i, chr;
  if (str.length === 0) return hash;
  for (i = 0; i < str.length; i++) {
    chr = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + chr;
    hash |= 0; // Convert to 32bit integer
  }
  return hash;
}