"""
Application script to allow the user to output graphs and data to 
a weekly report.

Fix: The Word templates uses the round(3) function, which does
not format to a fixed number of DP like "%0.4f". Formatted strings
seem to cause (charecter encoding) problems in Word though.

Author: sdo
Updated: 22/01/2018 (refactored to use for loops over network and indice names)
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
	"""Returns a string for the dollar penalty/incentive for a given 
	network's SAIDI or SAIFI"""
	if network == "TPCO":
		net = "TPC"
	elif network == "ELIN":
		net = "EIL"
	elif network == "OTPO":
		net = "OJV"
	ir = 0.5 * 0.01 * Constants.CC_Revenue_At_Risk.get(network) / (Constants.CC_Vals.get(network).get(saidi_saifi+"_CAP") - Constants.CC_Vals.get(network).get(saidi_saifi+"_TARGET") )
	full_year_forecast = float(params.get(net+"_"+saidi_saifi+"_UNPLANNED") + params.get(net+"_"+saidi_saifi+"_PLANNED")) - context.get(saidi_saifi+'_'+net+'_YTD_T') + Constants.CC_Vals.get(network).get(saidi_saifi+"_TARGET")
	#full_year_forecast = float(get_rtext(saidi_saifi+'_'+net+'_YTD')) - context.get(saidi_saifi+'_'+net+'_YTD_T') + Constants.CC_Vals.get(network).get(saidi_saifi+"_TARGET")
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
	
	# (optional) New formatting style (no $ sign)
	rewardstr = "%.2f" % (adjustedreward)
	return rewardstr


if __name__ == "__main__":
	networks = [["EIL", "ELIN"], ["TPC", "TPCO"], ["OJV", "OTPO"]]
	indices = ["SAIDI", "SAIFI"]
	colours = {}
	dictdir = os.path.join(Constants.FILE_DIRS.get("GENERAL"), "Stats")
	params = load_obj(dictdir, "paramsdict")

	context = {'TODAY_DATE'	: datetime.datetime.now().strftime("%A %d %B %Y"),
			}
	# The date of the results file was generted and saved
	context['SAIDI_SAIFI_DATE'] = params.get("EIL_DATE_END").strftime("%d/%m/%Y")

	# Create an object of the Word template document
	doc = DocxTemplate(os.path.join(Constants.FILE_DIRS.get("GENERAL"), "Templates", "Weekly Report Template.docx"))

	# Fill the template context dictionary
	for network in networks:
		for indice in indices:
			ref_string = "{0}_{1}".format(network[0], indice)

			# YTD Actuals - Rich Text (so round numeric values here instead of template)
			if params.get("{0}_{1}_UNPLANNED".format(network[0], indice)) + params.get("{0}_{1}_PLANNED".format(network[0], indice)) > params.get("{0}_CC_{1}_YTD".format(network[0], indice)):
				colours[ref_string] = "#FF0000"
			else:
				colours[ref_string] = "#3AB14D"
			context["{1}_{0}_YTD".format(network[0], indice)] = RichText("%.3f" % (params.get("{0}_{1}_UNPLANNED".format(network[0], indice)) + params.get("{0}_{1}_PLANNED".format(network[0], indice))), color=colours.get(ref_string), style="Report")

			# YTD Targets
			context["{0}_{1}_YTD_T".format(indice, network[0])] = params.get("{0}_CC_{1}_YTD".format(network[0], indice))

			# EOY Targets
			context["{0}_{1}_EOY_T".format(indice, network[0])] = Constants.CC_Vals.get(network[1]).get(indice+"_TARGET")
			
			# EOY Limits/Caps
			context["{0}_{1}_EOY_L".format(indice, network[0])] = Constants.CC_Vals.get(network[1]).get(indice+"_CAP")

			# Expected incentive/penalty - assumes SAIDI/SAIFI trend linearly at the target rate - Rich Text
			context["{0}_{1}_EIP".format(indice, network[0])] = RichText(
				(calcreward(indice, network[1])), color=colours.get(ref_string), style="Report")

			# Month to date interruptions
			context["{0}_{1}_MTD_UBV".format(indice, network[0])] = params.get("{0}_RAW_MONTH_NUM_MAJOR_EVENTS_{1}".format(network[0], indice))

			# Year to date interruptions
			context["{0}_{1}_YTD_UBV".format(indice, network[0])] = params.get("{0}_RAW_NUM_MAJOR_EVENTS_{1}".format(network[0], indice))

			# Add images of the charts to the document
			context["{0}_{1}_CHART".format(network[0], indice)] = InlineImage(doc, 
				os.path.join(Constants.FILE_DIRS.get("GENERAL"), "Stats/img/{0}_{1}.png".format(network[0], indice)), width=Mm(100), height=Mm(75))
		
		# Month to date interruptions	
		context["{0}_PLAN_MTD".format(network[0])] = params.get("{0}_RAW_MONTH_PLANNED".format(network[0]))
		context["{0}_UNPLAN_MTD".format(network[0])] = params.get("{0}_RAW_MONTH_UNPLANNED".format(network[0]))

		# Year to date interruptions
		context["{0}_PLAN_YTD".format(network[0])] = params.get("{0}_RAW_PLANNED".format(network[0]))
		context["{0}_UNPLAN_YTD".format(network[0])] = params.get("{0}_RAW_UNPLANNED".format(network[0]))

	# Populate the template params
	doc.render(context)

	# Save the rendered document
	try:
		doc.save(os.path.join(Constants.FILE_DIRS.get("GENERAL"), "Weekly Report.docx"))
	except IOError, e:
		# The document is probably open already
		print e
		print "The document is probably open, and can not be saved."

	print "Success!"
	time.sleep(5)