<?php
/**
 * The code is freely given by "big-pete" (his handle name in Joomla forums)
 * Original code is from URL: http://forum.joomla.org/viewtopic.php?t=433060
 * 
 * Requirement:
 * You must add a PHPExcel plugin in your Joomla!
 *   - (Current Stable Repo: https://github.com/PHPOffice/PHPExcel/tree/develop) as of June 2012
 *   - either you copy or symlink the Classes folder to your "libraries" folder
 *
 * Good luck and Happy Computing,
 * ~ Romack Natividad <romack@wizmediateam.com>
 */
defined('_JEXEC') or die();

// Include Joomla Classes
JLoader::import('joomla.filesystem.folder');

// Include PHPExcel classes
JLoader::import('phpexcel.PHPExcel');
JLoader::import('phpexcel.PHPExcel.IOFactory');

class JDocumentXLS extends JDocument
{
	private $_name = 'export';
	private $_phpexcel = null;

	public function __construct($options = array())
	{ 
		parent::__construct($options);
		$this->setMimeEncoding('application/vnd.ms-excel');
		$this->setType('xls');
		if(isset($options['name'])) $this->setName($options['name']);
	}

	/**
	 * Returns a PHPExcel object
	 * @param None
	 * @return Object
	*/
	public function &getPhpExcelObj()
	{ 
		if(!$this->_phpexcel) $this->_phpexcel = new PHPExcel();
		return $this->_phpexcel;
	}

	/**
	 * Renders the document
	 * @param boolean  $cache    If true, cache the output
	 * @param array   $params   Associative array of attributes
	 * @return String
	*/
	public function render($cache = false, $params = array())
	{ 
		// Write out response headers
		JResponse::setHeader('Pragma', 'public', true);
		JResponse::setHeader('Expires', '0', true);
		JResponse::setHeader('Cache-Control', 'must-revalidate, post-check=0, pre-check=0', true);
		JResponse::setHeader('Content-Type', 'application/force-download', true);
		JResponse::setHeader('Content-Type', 'application/octet-stream', true);
		JResponse::setHeader('Content-Type', 'application/download', true);
		JResponse::setHeader('Content-Transfer-Encoding', 'binary', true);

		// Set workbook properties to some defaults if not currently set
		// Currently all these properties are not set for the Excel5 (xls) writer but are here in case future versions do
		$objPhpExcel =& $this->getPhpExcelObj();
		$config = new JConfig();
		$workbook_properties = $objPhpExcel->getProperties();

		if(!$workbook_properties->getCategory()) $workbook_properties->setCategory('Exported Report From '. $config->sitename);
		if($workbook_properties->getCompany() == 'Microsoft Corporation' && $config->sitename) $workbook_properties->setCompany($config->sitename);
		if($workbook_properties->getCreator() == 'Unknown Creator' && $config->sitename) $workbook_properties->setCreator($config->sitename);
		if(!$workbook_properties->getDescription()) $workbook_properties->setDescription($this->getDescription());
		if(!$workbook_properties->getLastModifiedBy()) $workbook_properties->setLastModifiedBy($config->sitename);
		if(!$workbook_properties->getSubject()) $workbook_properties->setSubject($this->getTitle());
		if($workbook_properties->getTitle() == 'Untitled Spreadsheet' && $this->getTitle()) $workbook_properties->setTitle($this->getTitle());

		$objPhpExcel->setProperties($workbook_properties);

		// Get the Excel 5 type IO object to write out the binary document
		$objWriter = PHPExcel_IOFactory::createWriter($objPhpExcel, 'Excel5');

		if(JFolder::exists($config->tmp_path)) $objWriter->setTempDir($config->tmp_path);
		
		// Save the file to the PHP Output Stream and read the stream back in to set the buffer
		ob_start();

		$objWriter->save('php://output');
		$buffer = ob_get_contents();
		
		ob_end_clean();
		
		JResponse::setHeader('Content-disposition', 'attachment; filename="' . $this->getName() . '.' . $this->getType() . '"; size=' . strlen($buffer) . ';', true);
		return $buffer;
	}

	/**
	 * Initializes the PHPExcel object
	 * @param $objPhpExcel
	 * @return None
	*/
	public function setPhpExcelObj($objPhpExcel)
	{ 
		$this->_phpexcel = $objPhpExcel;
	}	

	/**
	 * Sets the current filename of the excel file
	 * @param $name
	 * @return None
	*/
	public function setName($name)
	{ 
		$this->_name = JFilterOutput::stringURLSafe($name);
	}
	
	/**
	 * Returns the current filename for the excel file
	 * @param None
	 * @return String
	*/
	public function getName()
	{ 
		return $this->_name;
	}
}
