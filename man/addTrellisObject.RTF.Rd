%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Do not modify this file since it was automatically generated from:
% 
%  rtf.R
% 
% by the Rdoc compiler part of the R.oo package.
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

 \name{addTrellisObject.RTF}
\alias{addTrellisObject.RTF}
\alias{RTF.addTrellisObject}
\alias{addTrellisObject.RTF}
\alias{addTrellisObject,RTF-method}


 \title{Insert a trellis plot object into the RTF document}

 \description{
	Insert a trellis plot object into the RTF document.
 }

 \usage{\method{addTrellisObject}{RTF}(this, trellis.object, width=3, height=0.3, res=300, ...)}

 \arguments{
 	\item{this}{An RTF object.}
 	\item{trellis.object}{The trellis plot object.}
 	\item{width}{Plot output width in inches.}
 	\item{height}{Plot output height in inches.}
 	\item{res}{Output resolution in dots per inch.}
 	\item{...}{Not used.}
 }

 \details{
 	Plots are added to the document as PNG objects.  Multipage trellis objects are
	spread across multiple pages in the RTF output file.
 }

 \examples{
 \dontrun{
 rtf<-RTF("output.doc",width=8.5,height=11,font.size=10,omi=c(1,1,1,1))
 if(require(lattice)) {
 	# multipage trellis object
 	p2<-xyplot(uptake ~ conc | Plant, CO2, layout = c(2,2))
 	addTrellisObject(rtf,trellis.object=p2,width=6,height=6,res=png.res)
 }
 done(rtf)
 }}

 \author{Michael E. Schaffer}

 \seealso{
   For more information see \code{\link{RTF}}.
 }
\keyword{internal}
\keyword{methods}
