%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Do not modify this file since it was automatically generated from:
% 
%  rtf.R
% 
% by the Rdoc compiler part of the R.oo package.
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

 \name{RTF}
\docType{class}
\alias{RTF}


 \title{The RTF class}

 \description{
	This is the class representing an RTF file output.

	Package:  rtf \cr
\bold{Class RTF}\cr

\code{\link[R.oo]{Object}}\cr
\code{~~|}\cr
\code{~~+--}\code{RTF}\cr

\bold{Directly known subclasses:}\cr
\cr

public static class \bold{RTF}\cr
extends \link[R.oo]{Object}\cr


 }

 \usage{RTF(file="", width=8.5, height=11, omi=c(1, 1, 1, 1), font.size=10)}

 \arguments{
 	\item{file}{The path of the output file.}
 	\item{width}{The width of the output page.}
 	\item{height}{The width of the output page.}
 	\item{omi}{A \code{\link[base]{vector}} representing the outer margins in inches (bottom, left, top, right).}
	\item{font.size}{Default font size for the document in points.}
 	\item{...}{Not used.}
 }

 \section{Fields and Methods}{
 	\bold{Methods:}\cr
\tabular{rll}{
 \tab \code{\link[rtf:addHeader.RTF]{addHeader}} \tab Insert a header into the RTF document.\cr
 \tab \code{\link[rtf:addNewLine.RTF]{addNewLine}} \tab Insert a new line into the RTF document.\cr
 \tab \code{\link[rtf:addPageBreak.RTF]{addPageBreak}} \tab Insert a page break into the RTF document optionally specifying new page settings.\cr
 \tab \code{\link[rtf:addParagraph.RTF]{addParagraph}} \tab Insert a paragraph into the RTF document.\cr
 \tab \code{\link[rtf:addPlot.RTF]{addPlot}} \tab Insert a plot into the RTF document.\cr
 \tab \code{\link[rtf:addSessionInfo.RTF]{addSessionInfo}} \tab Insert session information into the RTF document.\cr
 \tab \code{\link[rtf:addTable.RTF]{addTable}} \tab Insert a table into the RTF document.\cr
 \tab \code{\link[rtf:addText.RTF]{addText}} \tab Insert text into the RTF document.\cr
 \tab \code{\link[rtf:addTrellisObject.RTF]{addTrellisObject}} \tab Insert a trellis plot object into the RTF document.\cr
 \tab \code{\link[rtf:decreaseIndent.RTF]{decreaseIndent}} \tab Decrease RTF document indent.\cr
 \tab \code{\link[rtf:done.RTF]{done}} \tab Write and close the RTF output.\cr
 \tab \code{\link[rtf:endParagraph.RTF]{endParagraph}} \tab End a paragraph in the RTF document.\cr
 \tab \code{\link[rtf:increaseIndent.RTF]{increaseIndent}} \tab Increase RTF document indent.\cr
 \tab \code{\link[rtf:setFontSize.RTF]{setFontSize}} \tab Set RTF document font size.\cr
 \tab \code{\link[rtf:startParagraph.RTF]{startParagraph}} \tab Start a new paragraph in the RTF document.\cr
 \tab \code{\link[rtf:view.RTF]{view}} \tab View encoded RTF.\cr
}


 \bold{Methods inherited from Object}:\cr
[[, [[<-, $, $<-, as.character, attach, attachLocally, clearCache, clearLookupCache, clone, detach, equals, extend, finalize, gc, getEnvironment, getFieldModifier, getFieldModifiers, getFields, getInstantiationTime, getStaticInstance, hasField, hashCode, ll, load, objectSize, print, registerFinalizer, save

 }

 \examples{
 output<-"test_RTF-class.doc"
 png.res<-300

 rtf<-RTF(output,width=8.5,height=11,font.size=10,omi=c(1,1,1,1))
 addHeader(rtf,title="Test",subtitle="2011-08-15\n")
 addPlot(rtf,plot.fun=plot,width=6,height=6,res=300, iris[,1],iris[,2])

 # Try trellis plots
 if(require(lattice) & require(grid)) {
 	# single page trellis objects
 	addPageBreak(rtf, width=11,height=8.5,omi=c(0.5,0.5,0.5,0.5))

 	p <- histogram( ~ height | voice.part, data = singer, xlab="Height")
 	addTrellisObject(rtf,trellis.object=p,width=10,height=7.5,res=png.res)

 	p <- densityplot( ~ height | voice.part, data = singer, xlab = "Height")
 	addTrellisObject(rtf,trellis.object=p,width=10,height=7.5,res=png.res)

 	# multipage trellis object
 	p2<-xyplot(uptake ~ conc | Plant, CO2, layout = c(2,2))
 	addTrellisObject(rtf,trellis.object=p2,width=6,height=6,res=png.res)
 }

 addPageBreak(rtf, width=6,height=10,omi=c(0.5,0.5,0.5,0.5))
 addTable(rtf,as.data.frame(head(iris)),font.size=10,row.names=FALSE,NA.string="-")

 addSessionInfo(rtf)

 done(rtf)
 }

 \author{Michael E. Schaffer}

 \seealso{
 	For more information see \code{\link{RTF}}.
 }
\keyword{classes}
