library(whisker)

f1 <- function(b) { return(c(b,b*2,b*3,b*4,b*5));}
f2 <- function(b) { return(list(name=b, items=f1(b)));}
f3 <- function() { return(list(code=1234, objects = list(f2(1),f2(2))));}

data <- list(name="Markus", items=f3())

writeLines(whisker.render(readLines('test.html'), data), 'o.html')