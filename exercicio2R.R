###
## Programação em R
###

### Comparações
x <- 1
x > 10               # Falso
x <= 2               # Verdadeiro
x == 3               # Falso
x != 'casa'          # Verdadeiro
x > 0 & x < 4        # entre 0 e 4 (verdadeiro)
!(x > 0 & x < 4)     # não esta entre 0 e 4 (ou seja, a negação da verdade)??)
x=2
x < 2 | x > 2        # x menor que dois ou maior que 2...

is.numeric(10)       # Verdadeiro
is.character('a')    # Verdadeiro
is.logical('casa')   # Falso
is.null(NULL)        # Verdadeiro
is.finite(log(0))    # Falso
is.na(x)             # Falso

x <- 3
if (x > 5) {
  cat('x é um numero maior que 5\n')
} else {
  cat('Não é\n')
}

x <- -5
if (x > 5) {
  cat('x é um numero maior que 5\n')
} else if (x>=0 & x < 5) {
  cat('x esta entre 0 (inclusive) e 5\n')
} else if (x==5) {
  cat('x é exatamente igual a 5\n')
} else {
  cat('x é negativo\n')
}
