function testConfigLoad() {
  var config = loadConfig_();
  var keys = Object.keys(config).filter(function (k) { return k !== '_raw'; });
  Logger.log('Loaded %s config keys', keys.length);
  Logger.log(JSON.stringify(keys.slice(0, 20)));
  return keys.length;
}
