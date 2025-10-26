grammar MyGrammar;

start: expression EOF;

expression
    : additiveExpression
    ;

additiveExpression
    : multiplicativeExpression ( (ADD | SUBTRACT) multiplicativeExpression )*
    ;

multiplicativeExpression
    : powerExpression ( (MULTIPLY | DIVIDE) powerExpression )*
    ;

powerExpression
    : unaryExpression (POW unaryExpression)*
    ;

unaryExpression
    : (ADD | SUBTRACT)* primaryExpression
    ;

primaryExpression
    : NUMBER
    | CELL_REF
    | MIN '(' expression ',' expression ')'
    | MAX '(' expression ',' expression ')'
    | INC '(' expression ')'
    | DEC '(' expression ')'
    | '(' expression ')'
    ;

POW: '^';
ADD: '+';
SUBTRACT: '-';
MULTIPLY: '*';
DIVIDE: '/';
MIN: 'min';
MAX: 'max';
INC: 'inc';
DEC: 'dec';

CELL_REF: [A-Z]+ [0-9]+;
NUMBER: [0-9]+ ('.' [0-9]+)?;

WHITESPACE: [ \t\r\n]+ -> skip;