trim.leading <- function (x)  sub("^\\s+", "", x)

# returns string w/o trailing whitespace
trim.trailing <- function (x) sub("\\s+$", "", x)

# returns string w/o leading or trailing whitespace
trim <- function (x) gsub("^\\s+|\\s+$", "", x)

split_uneven_string = function(vector, max.length, sep.by = " "){
  matrix(unlist(lapply(vector,
                       function(x) {
                         temp = trim(unlist(strsplit(x, sep.by)))
                         c(temp, rep(NA, max.length - length(temp)))})), ncol = max.length, byrow = T)
}

rename_single = function(data, oldname, newname){
  
  colnames(data)[which(colnames(data) == oldname)] = newname
  #print(names(data))
  return(data)
}

asNumeric <- function(x) as.numeric(as.character(x))

