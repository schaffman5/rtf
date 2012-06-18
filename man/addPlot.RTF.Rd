%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Do not modify this file since it was automatically generated from:
% 
%  rtf.R
% 
% by the Rdoc compiler part of the R.oo package.
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

 \name{addPlot.RTF}
\alias{addPlot.RTF}
\alias{RTF.addPlot}
\alias{addPlot.RTF}
\alias{addPlot,RTF-method}


 \title{Insert a plot into the RTF document}

 \description{
	Insert a plot into the RTF document.
 }

 \usage{\method{addPlot}{RTF}(this, plot.fun=plot.fun, width=3, height=0.3, res=300, ...)}

 \arguments{
 	\item{this}{An RTF object.}
 	\item{plot.fun}{Plot function.}
 	\item{width}{Plot output width in inches.}
 	\item{height}{Plot output height in inches.}
 	\item{res}{Output resolution in dots per inch.}
 	\item{...}{Arguments for \code{plot.fun}.}
 }

 \details{
 	Plots are added to the document as PNG objects.  This function will work with all
   base graphics methods for plotting.  For more sophisticated plots, you may need to
   wrap your plot code into a function, and then pass a reference to that function to
   this method.  The parameters for the plot method would then get passed in as '...'
   above.

   To output a ggplot2 plot, simply assign the plot to a variable.  Then use 'print'
   as the plot function and pass in the plot variable assigned above.
 }

 \examples{
 rtf<-RTF("test_addPlot.doc",width=8.5,height=11,font.size=10,omi=c(1,1,1,1))
 addPlot(rtf,plot.fun=plot,width=6,height=6,res=300, iris[,1],iris[,2])
 done(rtf)
 }

 \author{Michael E. Schaffer}

 \seealso{
   For more information see \code{\link{RTF}}.
 }


\keyword{methods}
