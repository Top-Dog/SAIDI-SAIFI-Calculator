"""
Application script to allow the user to output graphs and data to 
a weekly report.

Fix: The word templates uses the round(3) function, which does
not format to a fixed number of DP like "%0.4f". Formatted strings
seem to cause (charecter encoding) problems in word though.
"""

import time, os, pickle, datetime
from SAIDISAIFI import Constants

from docxtpl import DocxTemplate, InlineImage, RichText
from docx.shared import Mm, Inches, Pt

def load_obj(filedir, filename):
	"""Loads a python object from a file"""
	with open(os.path.join(filedir, filename) + '.pkl', 'rb') as f:
		return pickle.load(f)

def cap_values(value, llimit, ulimit, revatrisk, calcreward):
	"""normalises the data to the cap or limit (collar)"""
	if value > ulimit:
		calcreward = -revatrisk / 2.0
	elif value < llimit:
		calcreward = revatrisk / 2.0
	return calcreward

def get_rtext(kwarg):
	"""Get the first value of the text in the richtext object,
	returns the key word arg if it is not rich text."""
	xml = context.get(kwarg).__str__()
	openingtag = """<w:t xml:space="preserve">"""
	try:
		l = xml.index(openingtag)
		h = xml.index("""</w:t>""")
	except ValueError:
		return kwarg
	return xml[l+len(openingtag):h]

def calcreward(saidi_saifi, network):
	"""Returns the dollar penalty/incentive for a given 
	network's SAIDI or SAIFI"""
	if network == "TPCO":
		net = "TPC"
	elif network == "ELIN":
		net = "EIL"
	elif network == "OTPO":
		net = "OJV"
	ir = 0.5 * 0.01 * Constants.CC_Revenue_At_Risk.get(network) / (Constants.CC_Vals.get(network).get(saidi_saifi+"_CAP") - Constants.CC_Vals.get(network).get(saidi_saifi+"_TARGET") )
	full_year_forecast = float(get_rtext(saidi_saifi+'_'+net+'_YTD')) - context.get(saidi_saifi+'_'+net+'_YTD_T') + Constants.CC_Vals.get(network).get(saidi_saifi+"_TARGET")
	#full_year_forecast = context.get(saidi_saifi+'_'+net+'_YTD') - context.get(saidi_saifi+'_'+net+'_YTD_T') + Constants.CC_Vals.get(network).get(saidi_saifi+"_TARGET")
	reward = ir * (Constants.CC_Vals.get(network).get(saidi_saifi+"_TARGET") - full_year_forecast)
	#adjustedreward = cap_values(context.get(saidi_saifi+'_'+net+'_YTD'), Constants.CC_Vals.get(network).get(saidi_saifi+"_COLLAR"), Constants.CC_Vals.get(network).get(saidi_saifi+"_CAP"), 0.01 * Constants.CC_Revenue_At_Risk.get(network), reward)
	adjustedreward = cap_values(full_year_forecast, Constants.CC_Vals.get(network).get(saidi_saifi+"_COLLAR"), Constants.CC_Vals.get(network).get(saidi_saifi+"_CAP"), 0.01 * Constants.CC_Revenue_At_Risk.get(network), reward)
	
	# Do the rounding for dollars here (instead of template, since rich text is text only)
	if adjustedreward >= 0:
		rewardstr = "$%.2f" % abs(adjustedreward)
	else:
		# Force the -ve sign to the left of the dollar sign
		rewardstr = "-$%.2f" % abs(adjustedreward)
	return rewardstr


if __name__ == "__main__":
	dictdir = os.path.join(Constants.FILE_DIRS.get("GENERAL"), "Stats")
	params = load_obj(dictdir, "paramsdict")

	context = {'TODAY_DATE'	: datetime.datetime.now().strftime("%A %d %B %Y"),
			}
	# The date of the results file was generted and saved
	context['SAIDI_SAIFI_DATE'] = params.get("EIL_DATE_END").strftime("%d/%m/%Y")

	# YTD Actuals - Rich Text (so round here instead of template)
	if params.get("EIL_SAIDI_UNPLANNED") + params.get("EIL_SAIDI_PLANNED") > params.get("EIL_CC_SAIDI_YTD"):
		colourEILSAIDI="#FF0000"
	else:
		colourEILSAIDI="#3AB14D"
	context['SAIDI_EIL_YTD'] = RichText("%.3f" % (params.get("EIL_SAIDI_UNPLANNED") + params.get("EIL_SAIDI_PLANNED")), color=colourEILSAIDI, style="Report")
	if params.get("TPC_SAIDI_UNPLANNED") + params.get("TPC_SAIDI_PLANNED") > params.get("TPC_CC_SAIDI_YTD"):
		colourTPCSAIDI="#FF0000"
	else:
		colourTPCSAIDI="#3AB14D"
	context['SAIDI_TPC_YTD'] = RichText("%.3f" % (params.get("TPC_SAIDI_UNPLANNED") + params.get("TPC_SAIDI_PLANNED")), color=colourTPCSAIDI, style="Report")
	if params.get("OJV_SAIDI_UNPLANNED") + params.get("OJV_SAIDI_PLANNED") > params.get("OJV_CC_SAIDI_YTD"):
		colourOJVSAIDI="#FF0000"
	else:
		colourOJVSAIDI="#3AB14D"
	context['SAIDI_OJV_YTD'] = RichText("%.3f" % (params.get("OJV_SAIDI_UNPLANNED") + params.get("OJV_SAIDI_PLANNED")), color=colourOJVSAIDI, style="Report")
	if params.get("EIL_SAIFI_UNPLANNED") + params.get("EIL_SAIFI_PLANNED") > params.get("EIL_CC_SAIFI_YTD"):
		colourEILSAIFI="#FF0000"
	else:
		colourEILSAIFI="#3AB14D"
	context['SAIFI_EIL_YTD'] = RichText("%.3f" % (params.get("EIL_SAIFI_UNPLANNED") + params.get("EIL_SAIFI_PLANNED")), color=colourEILSAIFI, style="Report")
	if params.get("TPC_SAIFI_UNPLANNED") + params.get("TPC_SAIFI_PLANNED") > params.get("TPC_CC_SAIFI_YTD"):
		colourTPCSAIFI="#FF0000"
	else:
		colourTPCSAIFI="#3AB14D"
	context['SAIFI_TPC_YTD'] = RichText("%.3f" % (params.get("TPC_SAIFI_UNPLANNED") + params.get("TPC_SAIFI_PLANNED")), color=colourTPCSAIFI, style="Report")
	if params.get("OJV_SAIFI_UNPLANNED") + params.get("OJV_SAIFI_PLANNED") > params.get("OJV_CC_SAIFI_YTD"):
		colourOJVSAIFI="#FF0000"
	else:
		colourOJVSAIFI="#3AB14D"
	context['SAIFI_OJV_YTD'] = RichText("%.3f" % (params.get("OJV_SAIFI_UNPLANNED") + params.get("OJV_SAIFI_PLANNED")), color=colourOJVSAIFI, style="Report")

	# YTD Targets
	context['SAIDI_EIL_YTD_T'] = params.get("EIL_CC_SAIDI_YTD")
	context['SAIDI_TPC_YTD_T'] = params.get("TPC_CC_SAIDI_YTD")
	context['SAIDI_OJV_YTD_T'] = params.get("OJV_CC_SAIDI_YTD")
	context['SAIFI_EIL_YTD_T'] = params.get("EIL_CC_SAIFI_YTD")
	context['SAIFI_TPC_YTD_T'] = params.get("TPC_CC_SAIFI_YTD")
	context['SAIFI_OJV_YTD_T'] = params.get("OJV_CC_SAIFI_YTD")

	# EOY Targets
	context['SAIDI_EIL_EOY_T'] = Constants.CC_Vals.get("ELIN").get("SAIDI_TARGET")
	context['SAIDI_TPC_EOY_T'] = Constants.CC_Vals.get("TPCO").get("SAIDI_TARGET")
	context['SAIDI_OJV_EOY_T'] = Constants.CC_Vals.get("OTPO").get("SAIDI_TARGET")
	context['SAIFI_EIL_EOY_T'] = Constants.CC_Vals.get("ELIN").get("SAIFI_TARGET")
	context['SAIFI_TPC_EOY_T'] = Constants.CC_Vals.get("TPCO").get("SAIFI_TARGET")
	context['SAIFI_OJV_EOY_T'] = Constants.CC_Vals.get("OTPO").get("SAIFI_TARGET")

	# EOY Limits/Caps
	context['SAIDI_EIL_EOY_L'] = Constants.CC_Vals.get("ELIN").get("SAIDI_CAP")
	context['SAIDI_TPC_EOY_L'] = Constants.CC_Vals.get("TPCO").get("SAIDI_CAP")
	context['SAIDI_OJV_EOY_L'] = Constants.CC_Vals.get("OTPO").get("SAIDI_CAP")
	context['SAIFI_EIL_EOY_L'] = Constants.CC_Vals.get("ELIN").get("SAIFI_CAP")
	context['SAIFI_TPC_EOY_L'] = Constants.CC_Vals.get("TPCO").get("SAIFI_CAP")
	context['SAIFI_OJV_EOY_L'] = Constants.CC_Vals.get("OTPO").get("SAIFI_CAP")

	# Expected incentive/penalty - assumes SAIDI/SAIFI trend linearly at the target rate - Rich Text
	context['SAIDI_EIL_EIP'] = RichText((calcreward("SAIDI", "ELIN")), color=colourEILSAIDI, style="Report")
	context['SAIDI_TPC_EIP'] = RichText((calcreward("SAIDI", "TPCO")), color=colourTPCSAIDI, style="Report")
	context['SAIDI_OJV_EIP'] = RichText((calcreward("SAIDI", "OTPO")), color=colourOJVSAIDI, style="Report")
	context['SAIFI_EIL_EIP'] = RichText((calcreward("SAIFI", "ELIN")), color=colourEILSAIFI, style="Report")
	context['SAIFI_TPC_EIP'] = RichText((calcreward("SAIFI", "TPCO")), color=colourTPCSAIFI, style="Report")
	context['SAIFI_OJV_EIP'] = RichText((calcreward("SAIFI", "OTPO")), color=colourOJVSAIFI, style="Report")

	# Create an object of the Word tempalte document
	doc = DocxTemplate(os.path.join(Constants.FILE_DIRS.get("GENERAL"), "Templates", "Weekly Report Template.docx"))

	# Add images of the charts to the document
	context['EIL_SAIDI_CHART'] = InlineImage(doc, os.path.join(Constants.FILE_DIRS.get("GENERAL"), 'Stats/img/EIL_SAIDI.png'), width=Mm(100), height=Mm(75))
	context['EIL_SAIFI_CHART'] = InlineImage(doc, os.path.join(Constants.FILE_DIRS.get("GENERAL"), 'Stats/img/EIL_SAIFI.png'), width=Mm(100), height=Mm(75))
	context['TPC_SAIDI_CHART'] = InlineImage(doc, os.path.join(Constants.FILE_DIRS.get("GENERAL"), 'Stats/img/TPC_SAIDI.png'), width=Mm(100), height=Mm(75))
	context['TPC_SAIFI_CHART'] = InlineImage(doc, os.path.join(Constants.FILE_DIRS.get("GENERAL"), 'Stats/img/TPC_SAIFI.png'), width=Mm(100), height=Mm(75))
	context['OJV_SAIDI_CHART'] = InlineImage(doc, os.path.join(Constants.FILE_DIRS.get("GENERAL"), 'Stats/img/OJV_SAIDI.png'), width=Mm(100), height=Mm(75))
	context['OJV_SAIFI_CHART'] = InlineImage(doc, os.path.join(Constants.FILE_DIRS.get("GENERAL"), 'Stats/img/OJV_SAIFI.png'), width=Mm(100), height=Mm(75))

	# Populate the template params
	doc.render(context)

	# Save the rendered document
	try:
		doc.save(os.path.join(Constants.FILE_DIRS.get("GENERAL"), "Weekly Report.docx"))
	except IOError, e:
		# The document is probably open already
		print e
		print "The docuemnt is probably open, and can not be saved."

	print "Success!"
	time.sleep(5)