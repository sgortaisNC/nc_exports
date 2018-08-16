<?php

namespace Drupal\export_xls\Form;

use Drupal\Core\Form\FormBase;
use Drupal\Core\Form\FormStateInterface;
use \Drupal\user\Entity\User;
use Drupal\Core\Datetime\DrupalDateTime;


use Drupal\export_xls\Controller\ExportXlsController;

class ExportAdminForm extends FormBase {
	/**
   * {@inheritdoc}
   */
	public function getFormId() {
		return 'export_xls_admin_form';
	}

	/**
	* {@inheritdoc}
	*/
	public function buildForm(array $form, FormStateInterface $form_state) {
		$form['type'] = array(
			'#type' => 'select',
			'#title' => 'Choisir le type d\'export',
			'#options' => [
				'actualite' => 'Actualités',
				'event' => 'Agendas',
				'app' => 'Appels à projet',
				'documents' => 'Documents',
				'offre_cooperation' => 'Offres de coopération',
				'projet' => 'Projets',
				'photo' => 'Photos',
				'video' => 'Vidéo',
				'annuaire' => 'Annuaire',
			],
	    );

		$form['date'] = array(
			'#type' => 'date',
			'#title' => "A partir du :",
			'#default_value' => '1970-01-01',
			'#description' => 'Sans impact pour "Annuaire"'
		);

		$form['submit'] = array(
			'#type' => 'submit',
			'#value' => 'Exporter',
		);

		return $form;
	}

	/**
	* {@inheritdoc}
	*/
	public function validateForm(array &$form, FormStateInterface $form_state) {

  	}

	/**
	* {@inheritdoc}
	*/
	public function submitForm(array &$form, FormStateInterface $form_state) {
		$export = new ExportXlsController;
		$retourImport = $export->export($form_state->getValues());
		drupal_set_message($retourImport);
  	}
}
