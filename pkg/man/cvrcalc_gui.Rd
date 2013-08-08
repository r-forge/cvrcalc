\name{cvrcalc_gui}
\alias{cvrcalc_gui}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
GUI for the cvrcalc package
}
\description{
A concise (1-5 lines) description of what the function does.
cvrclac_gui() es una función que llama a un entorno 
}
\usage{
cvrcalc_gui()
}
%- maybe also 'usage' for other objects documented here.
\details{
If necessary, more details than the description above
}
\value{
%%  ~Describe the value returned
%%  If it is a LIST, use
%%  \item{comp1 }{Description of 'comp1'}
%%  \item{comp2 }{Description of 'comp2'}
%% ...
}
\references{
%% ~put references to the literature/web site here ~
}
\author{
%%  ~~who you are~~
}
\note{
%%  ~~further notes~~
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\seealso{
%% ~~objects to See Also as \code{\link{help}}, ~~~
}
\examples{
##---- Should be DIRECTLY executable !! ----
##-- ==>  Define data, use random,
##--	or do  help(data=index)  for the standard data sets.

## The function is currently defined as
function () 
{
    modelos = c("Please, Select a model...", "Dorica", "Classic Framingham", 
        "Framingham-Wilson", "Regicor", "High Risk Score", "Low Risk Score")
    win = gwindow("CVR-CALC")
    group = ggroup(horizontal = FALSE, container = win)
    texto = glabel("A Cardiovascular Risk Calculator using estimation by Scores", 
        container = group, font.attr = list(style = "bold"))
    addSpring(group)
    addSpace(group, 15)
    modelo = gcombobox(modelos, container = group)
    addSpring(group)
    boton = gbutton("Run", container = group, handler = function(h, 
        ...) {
        eleccion = svalue(modelo)
        print(eleccion)
        if (eleccion == modelos[2]) 
            dorica()
        else if (eleccion == modelos[3]) 
            framingham_c()
        else if (eleccion == modelos[4]) 
            framingham_w()
        else if (eleccion == modelos[5]) 
            regicor()
        else if (eleccion == modelos[6]) 
            hrs()
        else if (eleccion == modelos[7]) 
            lrs()
        else print("Please, Select a model.")
    })
    mensaje_excel = "Please, Select the Excel file with the DATA to import..."
    mensaje_fin = "End of Process. Please, open de Excel file to view the results."
  }
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{ ~kwd1 }
\keyword{ ~kwd2 }% __ONLY ONE__ keyword per line
