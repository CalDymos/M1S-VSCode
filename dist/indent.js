var m1sindent = function m1sindent_(options){


  var
    parse = require('./parser.js'),
    formatting = require('./formatter.js');

  options = options || {};
  (function m1sindent_formatting(){
    var tparsed = parse(options);

    (function m1sindent_formatting_options(){
      options.tokens = tparsed.tokens || [];
      options.tokenTypes = tparsed.tokenTypes || [];
      options.level = options.level && /^\d+$/.test(options.level) ? parseInt(options.level) : 0;
      options.indentChar = options.indentChar ? options.indentChar.toString() : '  ';
      options.breakLineChar = options.breakLineChar ? options.breakLineChar.toString() : '\n';
      options.breakOnSeperator = options.breakOnSeperator === true || false;
      options.removeComments = options.removeComments === true || false;
    })();
  })();

  return formatting(options);
};

module.exports = m1sindent;
