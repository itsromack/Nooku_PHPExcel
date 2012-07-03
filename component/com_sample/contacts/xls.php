<?php
class ComMccedViewClassesXls extends ComDefaultViewXls
{
	protected function _initialize(KConfig $config)
	{
		$config->append(array(
			'title' => 'List of Contacts',
			'filename' => 'Contacts',
			'description' => 'List of Contacts',
			'headers' => array(
			    'Name',
				'Surname',
				'Email'
			)
		));

		parent::_initialize($config);
	}
}