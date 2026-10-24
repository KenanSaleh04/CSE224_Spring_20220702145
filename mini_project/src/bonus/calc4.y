%{
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <stdarg.h>
#include "calc4.h"

nodeType *opr(int oper, int nops, ...);
nodeType *id(int i);
nodeType *con(int value);
void freeNode(nodeType *p);
int ex(nodeType *p);
int yylex(void);

void yyerror(char *s);
extern int yylineno;
extern char *yytext;

symEntry sym[SYMMAX];
int nsym = 0;
%}

%union {
    int iValue;
    int sIndex;
    nodeType *nPtr;
};

%token <iValue> INTEGER
%token <sIndex> VARIABLE
%token WHILE DO FOR IF PRINT
%nonassoc IFX
%nonassoc ELSE

%left GE LE EQ NE '>' '<'
%left '+' '-'
%left '*' '/' '%'
%nonassoc UMINUS
%right '^'

%type <nPtr> stmt expr stmt_list assign

%%

program:
        function                { exit(0); }
        ;

function:
          function stmt         { ex($2); freeNode($2); }
        |
        ;

stmt:
          ';'                                       { $$ = opr(';', 2, NULL, NULL); }
        | expr ';'                                  { $$ = $1; }
        | PRINT expr ';'                            { $$ = opr(PRINT, 1, $2); }
        | assign ';'                                { $$ = $1; }
        | WHILE '(' expr ')' stmt                   { $$ = opr(WHILE, 2, $3, $5); }
        | DO stmt WHILE '(' expr ')' ';'            { $$ = opr(DO, 2, $2, $5); }
        | FOR '(' assign ';' expr ';' assign ')' stmt { $$ = opr(FOR, 4, $3, $5, $7, $9); }
        | IF '(' expr ')' stmt %prec IFX            { $$ = opr(IF, 2, $3, $5); }
        | IF '(' expr ')' stmt ELSE stmt            { $$ = opr(IF, 3, $3, $5, $7); }
        | '{' stmt_list '}'                         { $$ = $2; }
        ;

assign:
          VARIABLE '=' expr     { $$ = opr('=', 2, id($1), $3); }
        ;

stmt_list:
          stmt                  { $$ = $1; }
        | stmt_list stmt        { $$ = opr(';', 2, $1, $2); }
        ;

expr:
          INTEGER               { $$ = con($1); }
        | VARIABLE              { $$ = id($1); }
        | '-' expr %prec UMINUS { $$ = opr(UMINUS, 1, $2); }
        | expr '+' expr         { $$ = opr('+', 2, $1, $3); }
        | expr '-' expr         { $$ = opr('-', 2, $1, $3); }
        | expr '*' expr         { $$ = opr('*', 2, $1, $3); }
        | expr '/' expr         { $$ = opr('/', 2, $1, $3); }
        | expr '%' expr         { $$ = opr('%', 2, $1, $3); }
        | expr '^' expr         { $$ = opr('^', 2, $1, $3); }
        | expr '<' expr         { $$ = opr('<', 2, $1, $3); }
        | expr '>' expr         { $$ = opr('>', 2, $1, $3); }
        | expr GE expr          { $$ = opr(GE, 2, $1, $3); }
        | expr LE expr          { $$ = opr(LE, 2, $1, $3); }
        | expr NE expr          { $$ = opr(NE, 2, $1, $3); }
        | expr EQ expr          { $$ = opr(EQ, 2, $1, $3); }
        | '(' expr ')'          { $$ = $2; }
        ;

%%

int lookup(char *name) {
    int i;

    for (i = 0; i < nsym; i++)
        if (strcmp(sym[i].name, name) == 0)
            return i;
    if (nsym == SYMMAX) {
        yyerror("symbol table full");
        exit(1);
    }
    strncpy(sym[nsym].name, name, NAMEMAX - 1);
    sym[nsym].name[NAMEMAX - 1] = 0;
    sym[nsym].value = 0;
    return nsym++;
}

nodeType *con(int value) {
    nodeType *p;

    if ((p = malloc(sizeof(nodeType))) == NULL)
        yyerror("out of memory");

    p->type = typeCon;
    p->con.value = value;

    return p;
}

nodeType *id(int i) {
    nodeType *p;

    if ((p = malloc(sizeof(nodeType))) == NULL)
        yyerror("out of memory");

    p->type = typeId;
    p->id.i = i;

    return p;
}

nodeType *opr(int oper, int nops, ...) {
    va_list ap;
    nodeType *p;
    int i;

    if ((p = malloc(sizeof(nodeType) + (nops-1) * sizeof(nodeType *))) == NULL)
        yyerror("out of memory");

    p->type = typeOpr;
    p->opr.oper = oper;
    p->opr.nops = nops;
    va_start(ap, nops);
    for (i = 0; i < nops; i++)
        p->opr.op[i] = va_arg(ap, nodeType*);
    va_end(ap);
    return p;
}

void freeNode(nodeType *p) {
    int i;

    if (!p) return;
    if (p->type == typeOpr) {
        for (i = 0; i < p->opr.nops; i++)
            freeNode(p->opr.op[i]);
    }
    free (p);
}

void yyerror(char *s) {
    fprintf(stdout, "line %d: %s near '%s'\n", yylineno, s, yytext);
}

int main(void) {
    yyparse();
    return 0;
}
