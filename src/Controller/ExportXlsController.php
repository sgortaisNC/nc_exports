<?php

namespace Drupal\export_xls\Controller;

use Drupal\Core\Controller\ControllerBase;
use Drupal\Core\Datetime\DrupalDateTime;
use Drupal\Core\Link;
use Drupal\Core\Url;
use Drupal\field_collection\Entity\FieldCollectionItem;
use Drupal\file\Entity\File;
use Drupal\node\Entity\Node;
use Drupal\taxonomy\Entity\Term;
use Drupal\user\Entity\User;

global $tabFieldCollection;

/**
 * Controller routines for import_csv module routes.
 */
class ExportXlsController extends ControllerBase {

	public function export( $dataForm = null, $type = null, $uid = null ) {
		set_time_limit( 0 );
		$host = \Drupal::request()->getHost();

		$bundle = $dataForm['type'];
		$ban    = array(
			"builder",
			'feeds_item',
			'boolean'
		);
		if ( $dataForm['type'] != "annuaire" ) {
			$entity_type_id = 'node';

			$currentTime = \DateTime::createFromFormat( "Y-m-d H:i:s", $dataForm['date'] . ' 00:00:00' );
			$currentTime = $currentTime->format( 'U' );

			$nids  = \Drupal::entityQuery( 'node' )->condition( 'type', $bundle )->condition( "created", $currentTime, ">" )->execute();
			$nodes = \Drupal\node\Entity\Node::loadMultiple( $nids );
			foreach ( \Drupal::entityManager()->getFieldDefinitions( $entity_type_id, $bundle ) as $field_name => $field_definition ) {
				if ( ! empty( $field_definition->getTargetBundle() ) ) {
					$bundleFields[ $field_name ]['entity_ref'] = '';
					$bundleFields[ $field_name ]['type']       = $field_definition->getType();
					$bundleFields[ $field_name ]['label']      = $field_definition->getLabel();
					if ( $field_definition->getType() == "entity_reference" ) {
						$handler = $field_definition->get( "settings" );
						if ( $handler["handler"] == "default:taxonomy_term" ) {
							$bundleFields[ $field_name ]['entity_ref'] = "taxo";
						} else {
							$bundleFields[ $field_name ]['entity_ref'] = "node";
						}
					}
				}
			}
			$data = [];
			foreach ( $nodes as $cNode ) {
				$elem = [];
				foreach ( $bundleFields as $key => $dataField ) {
					$elem["Date de création"] = mb_convert_encoding( date( "d/m/Y H:i:s", $cNode->get( "created" )->getString() ), 'UTF-16', 'UTF-8' );
					if ( ! in_array( $dataField["type"], $ban ) ) {
						$cleElem = $dataField["label"];
						if ( $dataField["type"] == "entity_reference" ) {
							$datasEntRef = $cNode->get( $key )->getString();
							$datasEntRef = explode( ", ", $datasEntRef );
							$tmpEntRef   = [];
							if ( $dataField["entity_ref"] == "taxo" ) {
								if ( ! empty( $cNode->get( $key )->getString() ) ) {
									foreach ( $datasEntRef as $tid ) {
										if ( $term = \Drupal\taxonomy\Entity\Term::load( $tid ) ) {
											$tmpEntRef[] = $term->get( 'name' )->value;
										}
									}
								}
							} else {

								if ( ! empty( $cNode->get( $key )->getString() ) ) {
									foreach ( $datasEntRef as $nid ) {
										$nodeEntRef  = \Drupal\node\Entity\Node::load( $nid );
										$tmpEntRef[] = $nodeEntRef->getTitle();
									}
								}
							}
							$valueElem = mb_convert_encoding( implode( "#", $tmpEntRef ), 'UTF-16', 'UTF-8' );
						} else {
							if ( $dataField["type"] == "file" || $dataField["type"] == "image" ) {
								$valueStringified = $cNode->get( $key )->getString();
								$datasEntRef      = explode( ", ", $valueStringified );

								if ( $dataField["type"] == "image" ) {
									$datasEntRef = array_slice( $datasEntRef, 0, 1 );
								}
								$tmpEntRef = [];
								if ( ! empty( $cNode->get( $key )->getString() ) ) {

									foreach ( $datasEntRef as $fid ) {
										if ( is_numeric( $fid ) && (int) $fid > 1 ) {
											if ( $file = \Drupal\file\Entity\File::load( $fid ) ) {
												$tmpEntRef[] = str_replace( "public://", "https://" . $host . "/sites/default/files/", $file->getFileUri() );
											};
											unset( $file );
										}
									}
								}
								$valueElem = mb_convert_encoding( implode( ", ", $tmpEntRef ), 'UTF-16', 'UTF-8' );
							} else {
								$valueElem = mb_convert_encoding( $cNode->get( $key )->getString(), 'UTF-16', 'UTF-8' );
								if ( $dataField["type"] == "field_collection" ) {
									$valueStringified = $cNode->get( $key )->getString();
									$datasEntRef      = explode( ", ", $valueStringified );
									$tmpEntRef        = [];
									if ( ! empty( $cNode->get( $key )->getString() ) ) {

										foreach ( $datasEntRef as $fcid ) {
											if ( is_numeric( $fcid ) ) {
												if ( $fc = \Drupal\field_collection\Entity\FieldCollectionItem::load( $fcid ) ) {
													$tmpEntRef[] = /*$fc->get( "field_nom" )->getString()*/
														$fcid;
												};
											}
										}
									}
									$valueElem = mb_convert_encoding( implode( ", ", $tmpEntRef ), 'UTF-16', 'UTF-8' );

								}
							}
						}
						$elem[ $cleElem ] = $valueElem;


					}
				}
				$elemF = array();

				switch ( $bundle ) {
					case 'projet' :
						$elemDate    = explode( mb_convert_encoding( ", ", 'UTF-16', 'UTF-8' ), $elem["Dates"] );
						$elemdateDeb = $elemDate[0];
						if ( count( $elemDate ) > 1 ) {
							$elemdatefin = $elemDate[1];
						} else {
							$elemdatefin = '';
						}
						$elemThemesLeader = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Thèmes de coopération LEADER"] );
						$elemRubriques    = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Rubriques"] );
						$elemThemes       = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Thèmes"] );
						$elemThemesPEI    = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Thèmes PEI"] );
						$elemEchelons     = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Echelons géographiques"] );
						$elemZonePei      = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Zones géographiques (GO PEI)"] );
						$elemReco         = explode( mb_convert_encoding( ", full_html,  ", 'UTF-16', 'UTF-8' ), $elem["Recommandations et résultats du projet"] );
						$lastReco         = array_pop( $elemReco );
						$elemBody         = explode( mb_convert_encoding( ", full_html,  ", 'UTF-16', 'UTF-8' ), $elem["Description du projet et des activités"] );
						$lastBody         = array_pop( $elemReco );
						$elemContexte     = explode( mb_convert_encoding( ", full_html,  ", 'UTF-16', 'UTF-8' ), $elem["Contexte et objectifs du projet"] );
						$lastContexte     = array_pop( $elemContexte );

						$elemF = [
							"Identifiant (pour import)"                          => "",
							"Identifiant (ID)"                                   => $cNode->id(),
							"Titre du projet"                                    => $elem["Titre"],
							"Identifiant du projet (GO PEI)"                     => $elem["Identifiant de projet (GO PEI)"],
							"N° dossier administratif"                           => $elem["Numéro de dossier administratif"],
							"Auteur du texte"                                    => $elem["Auteur du texte"],
							"Date de mise à jour"                                => $elem["Date de création"],
							"Rubrique n°1"                                       => count( $elemRubriques ) > 0 ? $elemRubriques[0] : "",
							"Rubrique n°2"                                       => count( $elemRubriques ) > 1 ? $elemRubriques[1] : "",
							"Rubrique n°3"                                       => count( $elemRubriques ) > 2 ? $elemRubriques[2] : "",
							"Rubrique n°4"                                       => count( $elemRubriques ) > 3 ? $elemRubriques[3] : "",
							"Rubrique n°5"                                       => count( $elemRubriques ) > 4 ? $elemRubriques[4] : "",
							"Porteur de projet"                                  => $elem["Nom du porteur de projet"],
							"Nom du responsable de projet"                       => $elem["Nom du responsable de projet"],
							"Coordonnées GPS"                                    => $elem["Coordonnées GPS"],
							"Adresse du porteur de projet"                       => $elem["Adresse du porteur de projet"],
							"Courriel"                                           => $elem["Courriel"],
							"Téléphone"                                          => $elem["Téléphone"],
							"Site Internet du projet (lien hypertexte)"          => $elem["Site internet du projet"],
							"Thème n°1"                                          => count( $elemThemes ) > 0 ? $elemThemes[0] : "",
							"Thème n°2"                                          => count( $elemThemes ) > 1 ? $elemThemes[1] : "",
							"Thème n°3"                                          => count( $elemThemes ) > 2 ? $elemThemes[2] : "",
							"Thème n°4"                                          => count( $elemThemes ) > 3 ? $elemThemes[3] : "",
							"Thème n°5"                                          => count( $elemThemes ) > 4 ? $elemThemes[4] : "",
							"Echelon géographique n°1"                           => count( $elemEchelons ) > 0 ? $elemEchelons[0] : "",
							"Echelon géographique n°2"                           => count( $elemEchelons ) > 1 ? $elemEchelons[1] : "",
							"Echelon géographique n°3"                           => count( $elemEchelons ) > 2 ? $elemEchelons[2] : "",
							"Echelon géographique n°4"                           => count( $elemEchelons ) > 3 ? $elemEchelons[3] : "",
							"Principale zone géographique (GO PEI)"              => count( $elemZonePei ) > 0 ? $elemZonePei[0] : "",
							"Autre zone géographique n°1 (GO PEI)"               => count( $elemZonePei ) > 1 ? $elemZonePei[1] : "",
							"Autre zone géographique n°2 (GO PEI)"               => count( $elemZonePei ) > 2 ? $elemZonePei[2] : "",
							"Programme de développement rural"                   => $elem["Programme de développement rural"],
							"Fonds"                                              => $elem["Fonds"],
							"Groupe d'action locale LEADER"                      => $elem["Groupe d'action locale LEADER"],
							"Mesure du FEADER"                                   => $elem["Mesure du FEADER"],
							"Type de coopération LEADER"                         => $elem["Type de coopération LEADER"],
							"Langues parlées/comprises (COOP LEADER)"            => $elem["Langues parlées / comprises (COOP LEADER)"],
							"Thèmes coopération LEADER n°1"                      => count( $elemThemesLeader ) > 0 ? $elemThemesLeader[0] : "",
							"Thèmes coopération LEADER n°2"                      => count( $elemThemesLeader ) > 1 ? $elemThemesLeader[1] : "",
							"Thème PEI n°1"                                      => count( $elemThemesPEI ) > 0 ? $elemThemesPEI[0] : "",
							"Thème PEI n°2"                                      => count( $elemThemesPEI ) > 1 ? $elemThemesPEI[1] : "",
							"Thème PEI n°3"                                      => count( $elemThemesPEI ) > 2 ? $elemThemesPEI[2] : "",
							"Thème PEI n°4"                                      => count( $elemThemesPEI ) > 3 ? $elemThemesPEI[3] : "",
							"Thème PEI n°5"                                      => count( $elemThemesPEI ) > 4 ? $elemThemesPEI[4] : "",
							"Thème PEI n°6"                                      => count( $elemThemesPEI ) > 5 ? $elemThemesPEI[5] : "",
							"Thème PEI n°7"                                      => count( $elemThemesPEI ) > 6 ? $elemThemesPEI[6] : "",
							"Date de début du projet"                            => $elemdateDeb,
							"Date de fin du projet"                              => $elemdatefin,
							"Date d'approbation du projet de coopération LEADER" => $elem["Date d'approbation du projet de coopération LEADER"],
							"Statut"                                             => $elem["Statut"] == "0" ? mb_convert_encoding( "En cours", 'UTF-16', 'UTF-8' ) : mb_convert_encoding( "Terminé", 'UTF-16', 'UTF-8' ),
							"Coût total du projet"                               => $elem["Coût total"],
							"Contribution du FEADER"                             => $elem["Contribution du FEADER"],
							"Autres contributions publiques"                     => $elem["Autres contributions publiques"],
							"Financement privé"                                  => $elem["Financement privé"],
							"Châpo"                                              => $lastBody,
							"Contexte et objectifs du projet"                    => implode( mb_convert_encoding( ", ", 'UTF-16', 'UTF-8' ), $elemContexte ),
							"Description du projet et des activités"             => implode( mb_convert_encoding( ", ", 'UTF-16', 'UTF-8' ), $elemBody ),
							"Recommandations et résultats du projet"             => implode( mb_convert_encoding( ", ", 'UTF-16', 'UTF-8' ), $elemReco ),
							"Nom du témoin"                                      => $elem["Nom du témoin"],
							"Fonction du témoin"                                 => $elem["Fonction du témoin"],
							"Témoignage"                                         => $elem["Message du témoin"],
							"Type de la tête de gondole"                         => $elem["Tête de gondole"],
							"Visuel statique"                                    => $elem["Visuel"],
							"Image du slider n°1"                                => '',
							"Image du slider n°2"                                => '',
							"Image du slider n°3"                                => '',
							"Vidéo du slider n°1"                                => '',
							"Vidéo du slider n°2"                                => '',
							"Vidéo du slider n°3"                                => '',
						];

						for ( $npart = 1; $npart <= 50; $npart ++ ) {
							$tabPart = explode( mb_convert_encoding( ", ", 'UTF-16', 'UTF-8' ), $elem["Partenaires"] );
							if ( $npart <= count($tabPart) ) {
								$fcid = $tabPart[ $npart - 1 ];
								if ( $fc = \Drupal\field_collection\Entity\FieldCollectionItem::load( (int) mb_convert_encoding( $fcid, 'UTF-8', 'UTF-16' ) ) ) {
									$partnerPays = $partnerPDR = $partnerType =  "";
									$elemF[ "Nom du partenaire n°" . $npart ]      = mb_convert_encoding( $fc->get( "field_nom" )->getString(), 'UTF-16', "UTF-8" );
									$elemF[ "Adresse - partenaire n°" . $npart ]   = mb_convert_encoding( $fc->get( "field_adresse" )->getString(), 'UTF-16', "UTF-8" );
									$elemF[ "Courriel - partenaire n°" . $npart ]  = mb_convert_encoding( $fc->get( "field_courriel" )->getString(), 'UTF-16', "UTF-8" );
									$elemF[ "Téléphone - partenaire n°" . $npart ] = mb_convert_encoding( $fc->get( "field_telephone" )->getString(), 'UTF-16', "UTF-8" );

									if ( $term = \Drupal\taxonomy\Entity\Term::load( (int) $fc->get( "field_pays" )->getString() ) ) {
										$partnerPays = $term->get( 'name' )->value;
									}
									$elemF[ "Pays concernés (COOP LEADER) - partenaire n°" . $npart ] = $partnerPays;

									if ( $term = \Drupal\taxonomy\Entity\Term::load( (int) $fc->get( "field_pdr" )->getString() ) ) {
										$partnerPDR = $term->get( 'name' )->value;
									}
									$elemF[ "Programme de développement rural - partenaire n°" . $npart ] = $partnerPDR;

									if ( $term = \Drupal\taxonomy\Entity\Term::load( (int) $fc->get( "field_type" )->getString() ) ) {
										$partnerType = $term->get( 'name' )->value;
									}
									$elemF[ "Type de partenaire (GO PEI) - partenaire n°" . $npart ] = $partnerType;
								};
							}
						}
						break;
					case 'actualite':
						$elemRubriques = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Rubriques"] );
						$elemFichiers  = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Fichiers"] );
						$elemThemes    = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Thèmes"] );
						$elemEchelons  = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Echelons géographiques"] );
						$elembody      = str_replace( mb_convert_encoding( ', full_html', 'UTF-16', 'UTF-8' ), "CHAPO", $elem["Contenu"] );
						$elemF         = [
							"Identifiant (ID)"           => $cNode->id(),
							"Titre"                      => $elem["Titre"],
							"Rubrique n°1"               => count( $elemRubriques ) > 0 ? $elemRubriques[0] : "",
							"Rubrique n°2"               => count( $elemRubriques ) > 1 ? $elemRubriques[1] : "",
							"Rubrique n°3"               => count( $elemRubriques ) > 2 ? $elemRubriques[2] : "",
							"Rubrique n°4"               => count( $elemRubriques ) > 3 ? $elemRubriques[3] : "",
							"Rubrique n°5"               => count( $elemRubriques ) > 4 ? $elemRubriques[4] : "",
							"Thème n°1"                  => count( $elemThemes ) > 0 ? $elemThemes[0] : "",
							"Thème n°2"                  => count( $elemThemes ) > 1 ? $elemThemes[1] : "",
							"Thème n°3"                  => count( $elemThemes ) > 2 ? $elemThemes[2] : "",
							"Thème n°4"                  => count( $elemThemes ) > 3 ? $elemThemes[3] : "",
							"Thème n°5"                  => count( $elemThemes ) > 4 ? $elemThemes[4] : "",
							"Echelon géographique n°1"   => count( $elemEchelons ) > 0 ? $elemEchelons[0] : "",
							"Echelon géographique n°2"   => count( $elemEchelons ) > 1 ? $elemEchelons[1] : "",
							"Echelon géographique n°3"   => count( $elemEchelons ) > 2 ? $elemEchelons[2] : "",
							"Fonds"                      => $elem["Fonds"],
							"Date de publication"        => $elem["Publication le :"],
							"Source de l'article"        => str_replace( mb_convert_encoding( ', restricted_html', 'UTF-16', 'UTF-8' ), "", str_replace( mb_convert_encoding( ', full_html', 'UTF-16', 'UTF-8' ), "", $elem["Source de l'article"] ) ),
							"Liens externes"             => str_replace( mb_convert_encoding( ', full_html', 'UTF-16', 'UTF-8' ), "", $elem["Liens externes"] ),
							"Chapô"                      => explode( 'CHAPO', $elembody )[1],
							"Contenu"                    => explode( 'CHAPO', $elembody )[0],
							"Type de la tête de gondole" => $elem["Tête de gondole"],
							"Visuel statique"            => empty( $elem["Image"] ) ? $elem["Image de la photothèque"] : $elem["Image"],
							"Image du slider n°1"        => $elem["Tête de gondole"],
							"Image du slider n°2"        => $elem["Tête de gondole"],
							"Image du slider n°3"        => $elem["Tête de gondole"],
							"Projet du slider n°1"       => $elem["Tête de gondole"],
							"Projet du slider n°2"       => $elem["Tête de gondole"],
							"Projet du slider n°3"       => $elem["Tête de gondole"],
						];
						break;
					case 'app':
						$elemDate    = explode( mb_convert_encoding( ", ", 'UTF-16', 'UTF-8' ), $elem["Dates"] );
						$elemdateDeb = $elemDate[0];
						if ( count( $elemDate ) > 1 ) {
							$elemdatefin = $elemDate[1];
						} else {
							$elemdatefin = '';
						}

						$elemRubriques = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Rubriques"] );
						$elemFichiers  = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Fichiers"] );
						$elemThemes    = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Thèmes"] );
						$elemEchelons  = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Echelons géographiques"] );
						$elembody      = str_replace( mb_convert_encoding( ', full_html', 'UTF-16', 'UTF-8' ), "", $elem["Plus d'informations"] );
						$elemF         = [
							"Identifiant (ID)"                 => $cNode->id(),
							"Intitulé de l'appel à projet"     => $elem["Titre"],
							"Rubrique n°1"                     => count( $elemRubriques ) > 0 ? $elemRubriques[0] : "",
							"Rubrique n°2"                     => count( $elemRubriques ) > 1 ? $elemRubriques[1] : "",
							"Rubrique n°3"                     => count( $elemRubriques ) > 2 ? $elemRubriques[2] : "",
							"Rubrique n°4"                     => count( $elemRubriques ) > 3 ? $elemRubriques[3] : "",
							"Rubrique n°5"                     => count( $elemRubriques ) > 4 ? $elemRubriques[4] : "",
							"Thème n°1"                        => count( $elemThemes ) > 0 ? $elemThemes[0] : "",
							"Thème n°2"                        => count( $elemThemes ) > 1 ? $elemThemes[1] : "",
							"Thème n°3"                        => count( $elemThemes ) > 2 ? $elemThemes[2] : "",
							"Thème n°4"                        => count( $elemThemes ) > 3 ? $elemThemes[3] : "",
							"Thème n°5"                        => count( $elemThemes ) > 4 ? $elemThemes[4] : "",
							"Organisme porteur"                => str_replace( mb_convert_encoding( ', full_html', 'UTF-16', 'UTF-8' ), "", $elem["Organisme porteur"] ),
							"Echelon géographique n°1"         => count( $elemEchelons ) > 0 ? $elemEchelons[0] : "",
							"Echelon géographique n°2"         => count( $elemEchelons ) > 1 ? $elemEchelons[1] : "",
							"Echelon géographique n°3"         => count( $elemEchelons ) > 2 ? $elemEchelons[2] : "",
							"Programme de développement rural" => $elem["Programme de développement rural"],
							"Mesure du FEADER concernée"       => $elem["Mesure du FEADER concernée"],
							"Fonds"                            => $elem["Fonds"],
							"Date d'ouverture"                 => $elemdateDeb,
							"Date de clôture"                  => $elemdatefin,
							"Plus d'informations"              => $elembody,
						];
						break;
					case 'documents':
						$elemRubriques = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Rubriques"] );
						$elemFichiers  = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Fichiers"] );
						$elemThemes    = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Thèmes"] );
						$elemEchelons  = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Echelons géographiques"] );

						$elemF = [
							"Identifiant (ID)"               => $cNode->id(),
							"Titre du projet"                => $elem["Titre"],
							"Rédacteur"                      => $elem["Rédacteur"],
							"Rubrique n°1"                   => count( $elemRubriques ) > 0 ? $elemRubriques[0] : "",
							"Rubrique n°2"                   => count( $elemRubriques ) > 1 ? $elemRubriques[1] : "",
							"Rubrique n°3"                   => count( $elemRubriques ) > 2 ? $elemRubriques[2] : "",
							"Rubrique n°4"                   => count( $elemRubriques ) > 3 ? $elemRubriques[3] : "",
							"Rubrique n°5"                   => count( $elemRubriques ) > 4 ? $elemRubriques[4] : "",
							"Thème n°1"                      => count( $elemThemes ) > 0 ? $elemThemes[0] : "",
							"Thème n°2"                      => count( $elemThemes ) > 1 ? $elemThemes[1] : "",
							"Thème n°3"                      => count( $elemThemes ) > 2 ? $elemThemes[2] : "",
							"Thème n°4"                      => count( $elemThemes ) > 3 ? $elemThemes[3] : "",
							"Thème n°5"                      => count( $elemThemes ) > 4 ? $elemThemes[4] : "",
							"Thème n°6"                      => count( $elemThemes ) > 5 ? $elemThemes[5] : "",
							"Thème n°7"                      => count( $elemThemes ) > 6 ? $elemThemes[6] : "",
							"Thème n°8"                      => count( $elemThemes ) > 7 ? $elemThemes[7] : "",
							"Echelon géographique n°1"       => count( $elemEchelons ) > 0 ? $elemEchelons[0] : "",
							"Echelon géographique n°2"       => count( $elemEchelons ) > 1 ? $elemEchelons[1] : "",
							"Echelon géographique n°3"       => count( $elemEchelons ) > 2 ? $elemEchelons[2] : "",
							"Echelon géographique n°4"       => count( $elemEchelons ) > 3 ? $elemEchelons[3] : "",
							"Fonds"                          => $elem["Fonds"],
							"Type de document"               => $elem["Type de document"],
							"Année de parution ou d'édition" => $elem["Année de parution ou d'édition"],
							"Contenu / descriptif"           => str_replace( mb_convert_encoding( ', full_html', 'UTF-16', 'UTF-8' ), "", $elem["Contenu / descriptif"] ),
							"Fichier n°1"                    => count( $elemFichiers ) > 0 ? $elemFichiers[0] : "",
							"Fichier n°2"                    => count( $elemFichiers ) > 1 ? $elemFichiers[1] : "",
							"Fichier n°3"                    => count( $elemFichiers ) > 2 ? $elemFichiers[2] : "",
							"Fichier n°4"                    => count( $elemFichiers ) > 3 ? $elemFichiers[3] : "",
							"Fichier n°5"                    => count( $elemFichiers ) > 4 ? $elemFichiers[4] : "",
							"Fichier n°6"                    => count( $elemFichiers ) > 5 ? $elemFichiers[5] : "",
							"Fichier n°7"                    => count( $elemFichiers ) > 6 ? $elemFichiers[6] : "",
							"Fichier n°8"                    => count( $elemFichiers ) > 7 ? $elemFichiers[7] : "",
							"Fichier n°9"                    => count( $elemFichiers ) > 8 ? $elemFichiers[8] : "",
							"Fichier n°10"                   => count( $elemFichiers ) > 9 ? $elemFichiers[9] : "",
							"Fichier n°11"                   => count( $elemFichiers ) > 10 ? $elemFichiers[10] : "",
							"Fichier n°12"                   => count( $elemFichiers ) > 11 ? $elemFichiers[11] : "",
							"Fichier n°13"                   => count( $elemFichiers ) > 12 ? $elemFichiers[12] : "",
							"Fichier n°14"                   => count( $elemFichiers ) > 13 ? $elemFichiers[13] : "",
							"Fichier n°15"                   => count( $elemFichiers ) > 14 ? $elemFichiers[14] : "",
						];

						break;
					case 'event':

						$elemRubriques = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Rubriques"] );
						$elemThemes    = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Thèmes"] );
						$elemEchelons  = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Echelons géographiques"] );

						$elemDate    = explode( mb_convert_encoding( ", ", 'UTF-16', 'UTF-8' ), $elem["Dates"] );
						$elemdateDeb = $elemDate[0];
						if ( count( $elemDate ) > 1 ) {
							$elemdatefin = $elemDate[1];
						} else {
							$elemdatefin = '';
						}

						$elembody = str_replace( mb_convert_encoding( ', full_html', 'UTF-16', 'UTF-8' ), "CHAPO", $elem["Contenu"] );


						$elemF = [
							"Identifiant (ID)"               => $cNode->id(),
							"Titre"                          => $elem["Titre"],
							"Rubrique n°1"                   => count( $elemRubriques ) > 0 ? $elemRubriques[0] : "",
							"Rubrique n°2"                   => count( $elemRubriques ) > 1 ? $elemRubriques[1] : "",
							"Rubrique n°3"                   => count( $elemRubriques ) > 2 ? $elemRubriques[2] : "",
							"Rubrique n°4"                   => count( $elemRubriques ) > 3 ? $elemRubriques[3] : "",
							"Rubrique n°5"                   => count( $elemRubriques ) > 4 ? $elemRubriques[4] : "",
							"Thème n°1"                      => count( $elemThemes ) > 0 ? $elemThemes[0] : "",
							"Thème n°2"                      => count( $elemThemes ) > 1 ? $elemThemes[1] : "",
							"Thème n°3"                      => count( $elemThemes ) > 2 ? $elemThemes[2] : "",
							"Thème n°4"                      => count( $elemThemes ) > 3 ? $elemThemes[3] : "",
							"Thème n°5"                      => count( $elemThemes ) > 4 ? $elemThemes[4] : "",
							"Echelon géographique n°1"       => count( $elemEchelons ) > 0 ? $elemEchelons[0] : "",
							"Echelon géographique n°2"       => count( $elemEchelons ) > 1 ? $elemEchelons[1] : "",
							"Echelon géographique n°3"       => count( $elemEchelons ) > 2 ? $elemEchelons[2] : "",
							"Fonds"                          => $elem["Fonds"],
							"Type d'évènement"               => $elem["Catégorie de l'événement"],
							"Date de début (ou date unique)" => $elemdateDeb,
							"Date de fin"                    => $elemdatefin,
							"Chapô"                          => explode( 'CHAPO', $elembody )[1],
							"Contenu"                        => explode( 'CHAPO', $elembody )[0],
							"Adresse de l'événement"         => $elem["Adresse postale"],
							"Fichier n°1 (programme)"        => $elem["Fichiers"],
							"En savoir plus (lien externe)"  => $elem["En savoir plus"],
							"Type de la tête de gondole"     => $elem["Tête de gondole"],
							"Visuel statique"                => $elem["Vignette"],
							"Image du slider n°1"            => $elem["Tête de gondole"],
							"Image du slider n°2"            => $elem["Tête de gondole"],
							"Image du slider n°3"            => $elem["Tête de gondole"],
							"Projet du slider n°1"           => $elem["Tête de gondole"],
							"Projet du slider n°2"           => $elem["Tête de gondole"],
							"Projet du slider n°3"           => $elem["Tête de gondole"],
						];
						break;
					case 'photo':
						$elemRubriques = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Rubriques"] );
						$elemThemes    = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Thèmes"] );
						$elemEchelons  = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Echelons géographiques"] );

						$elemF = [
							"Identifiant (pour import)"   => "",
							"Identifiant (ID)"            => $cNode->id(),
							"Titre du projet"             => $elem["Titre"],
							"Rubrique n°1"                => count( $elemRubriques ) > 0 ? $elemRubriques[0] : "",
							"Rubrique n°2"                => count( $elemRubriques ) > 1 ? $elemRubriques[1] : "",
							"Rubrique n°3"                => count( $elemRubriques ) > 2 ? $elemRubriques[2] : "",
							"Rubrique n°4"                => count( $elemRubriques ) > 3 ? $elemRubriques[3] : "",
							"Rubrique n°5"                => count( $elemRubriques ) > 4 ? $elemRubriques[4] : "",
							"Thème n°1"                   => count( $elemThemes ) > 0 ? $elemThemes[0] : "",
							"Thème n°2"                   => count( $elemThemes ) > 1 ? $elemThemes[1] : "",
							"Thème n°3"                   => count( $elemThemes ) > 2 ? $elemThemes[2] : "",
							"Thème n°4"                   => count( $elemThemes ) > 3 ? $elemThemes[3] : "",
							"Thème n°5"                   => count( $elemThemes ) > 4 ? $elemThemes[4] : "",
							"Echelon géographique n°1"    => count( $elemEchelons ) > 0 ? $elemEchelons[0] : "",
							"Echelon géographique n°2"    => count( $elemEchelons ) > 1 ? $elemEchelons[1] : "",
							"Echelon géographique n°3"    => count( $elemEchelons ) > 2 ? $elemEchelons[2] : "",
							"Fonds"                       => $elem["Fonds"],
							"Auteur (de la prise de vue)" => $elem["Auteur de la prise de vue"],
							"Copyright (propriétaire)"    => $elem["Copyright"],
							"Date de la prise de vue"     => $elem["Date de la prise de vue"],
							"Descriptif"                  => str_replace( mb_convert_encoding( ', restricted_html', 'UTF-16', 'UTF-8' ), "", $elem["Descriptif"] ),
							"Visuel"                      => $elem["Visuel"],
						];

						break;
					case 'video':
						$elemRubriques = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Rubriques"] );
						$elemThemes    = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Thèmes"] );
						$elemEchelons  = explode( mb_convert_encoding( "#", 'UTF-16', 'UTF-8' ), $elem["Echelons géographiques"] );
						$elemF         = [
							"Identifiant (pour import)" => "",
							"Identifiant (ID)"          => $cNode->id(),
							"Titre du projet"           => $elem["Titre"],
							"Rubrique n°1"              => count( $elemRubriques ) > 0 ? $elemRubriques[0] : "",
							"Rubrique n°2"              => count( $elemRubriques ) > 1 ? $elemRubriques[1] : "",
							"Rubrique n°3"              => count( $elemRubriques ) > 2 ? $elemRubriques[2] : "",
							"Thème n°1"                 => count( $elemThemes ) > 0 ? $elemThemes[0] : "",
							"Thème n°2"                 => count( $elemThemes ) > 1 ? $elemThemes[1] : "",
							"Thème n°3"                 => count( $elemThemes ) > 2 ? $elemThemes[2] : "",
							"Thème n°4"                 => count( $elemThemes ) > 3 ? $elemThemes[3] : "",
							"Thème n°5"                 => count( $elemThemes ) > 4 ? $elemThemes[4] : "",
							"Thème n°6"                 => count( $elemThemes ) > 5 ? $elemThemes[5] : "",
							"Thème n°7"                 => count( $elemThemes ) > 6 ? $elemThemes[6] : "",
							"Thème n°8"                 => count( $elemThemes ) > 7 ? $elemThemes[7] : "",
							"Echelon géographique n°1"  => count( $elemEchelons ) > 0 ? $elemEchelons[0] : "",
							"Echelon géographique n°2"  => count( $elemEchelons ) > 1 ? $elemEchelons[1] : "",
							"Echelon géographique n°3"  => count( $elemEchelons ) > 2 ? $elemEchelons[2] : "",
							"Fonds"                     => $elem["Fonds"],
							"Auteur ou source"          => $elem["Auteur ou source"],
							"Date de la vidéo"          => $elem["Année"],
							"Contenu"                   => str_replace( mb_convert_encoding( ', restricted_html', 'UTF-16', 'UTF-8' ), "", $elem["Contenu"] ),
							"Projet rattaché"           => $elem["Projet rattaché"],
							"Vidéo Youtube"             => explode( mb_convert_encoding( ",", 'UTF-16', 'UTF-8' ), $elem["Vidéo"] )[0],
						];
						break;
					default:
						$elemF = $elem;
						break;
				}
				$data[] = $elemF;
			}

			$header = [ mb_convert_encoding( "Date de création", 'UTF-16', 'UTF-8' ) ];
			foreach ( $bundleFields as $key => $dataField ) {
				if ( ! in_array( $dataField["type"], $ban ) ) {
					switch ( $bundle ) {
						case 'projet' :
							$header   = array( mb_convert_encoding( "Identifiant (pour import)", 'UTF-16', 'UTF-8' ) );
							$header[] = mb_convert_encoding( "Identifiant (ID)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Titre du projet", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Identifiant du projet (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "N° dossier administratif", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Auteur du texte", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Date de mise à jour", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°4", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°5", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Porteur de projet", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du responsable de projet", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Coordonnées GPS", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse du porteur de projet", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Site Internet du projet (lien hypertexte)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°4", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°5", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°4", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Principale zone géographique (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Autre zone géographique n°1 (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Autre zone géographique n°2 (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fonds", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Groupe d'action locale LEADER", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Mesure du FEADER'", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de coopération LEADER", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Langues parlées/comprises (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thèmes coopération LEADER n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thèmes coopération LEADER n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème PEI n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème PEI n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème PEI n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème PEI n°4", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème PEI n°5", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème PEI n°6", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème PEI n°7", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Date de début du projet", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Date de fin du projet", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Date d'approbation du projet de coopération LEADER", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Statut", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Coût total du projet", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Contribution du FEADER", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Autres contributions publiques", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Financement privé", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Châpo", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Contexte et objectifs du projet", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Description du projet et des activités", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Recommandations et résultats du projet", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du témoin", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fonction du témoin", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Témoignage", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de la tête de gondole", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Visuel statique", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Image du slider n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Image du slider n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Image du slider n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Vidéo du slider n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Vidéo du slider n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Vidéo du slider n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°4", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°5", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°6", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°7", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°8", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°9", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°10", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°11", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°12", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°13", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°14", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°15", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°16", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°17", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°18", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°19", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°20", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°21", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°22", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°23", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°24", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°25", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°26", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°27", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°28", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°29", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°30", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°31", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°32", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°33", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°34", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°35", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°36", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°37", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°38", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°39", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°40", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°41", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°42", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°43", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°44", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°45", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°46", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°47", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°48", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°49", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Nom du partenaire n°50", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Courriel", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Téléphone", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Pays concernés (COOP LEADER)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de partenaire (GO PEI)", 'UTF-16', 'UTF-8' );
							break;
						case 'actualite':
							$header   = array( mb_convert_encoding( "Identifiant (ID)", 'UTF-16', 'UTF-8' ) );
							$header[] = mb_convert_encoding( "Titre", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°4", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°5", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°4", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°5", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fonds", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Date de publication", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Source de l'article", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Liens externes", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Chapô", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Contenu ", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de la tête de gondole", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Visuel statique", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Image du slider n°1 ", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Image du slider n°2 ", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Image du slider n°3 ", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Projet du slider n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Projet du slider n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Projet du slider n°3", 'UTF-16', 'UTF-8' );
							break;
						case 'app':
							$header   = array( mb_convert_encoding( "Identifiant (ID)", 'UTF-16', 'UTF-8' ) );
							$header[] = mb_convert_encoding( "Intitulé de l'appel à projet", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°4", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°5", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°4", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°5", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Organisme porteur", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Programme de développement rural", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Mesure du FEADER concernée", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fonds", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Date d'ouverture", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Date de clôture", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Plus d'informations", 'UTF-16', 'UTF-8' );

							break;
						case 'documents':
							$header   = array( mb_convert_encoding( "Identifiant (ID)", 'UTF-16', 'UTF-8' ) );
							$header[] = mb_convert_encoding( "Titre", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rédacteur", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°4", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°5", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°4", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°5", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°6", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°7", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°8", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fonds", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de document", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Année de parution ou d'édition", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Contenu / descriptif", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fichier n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fichier n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fichier n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fichier n°4", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fichier n°5", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fichier n°6", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fichier n°7", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fichier n°8", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fichier n°9", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fichier n°10", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fichier n°11", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fichier n°12", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fichier n°13", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fichier n°14", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fichier n°15", 'UTF-16', 'UTF-8' );
							break;
						case 'event':
							$header   = array( mb_convert_encoding( "Identifiant (ID)", 'UTF-16', 'UTF-8' ) );
							$header[] = mb_convert_encoding( "Titre", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°4", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°5", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°4", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°5", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fonds", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type d'évènement", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Date de début (ou date unique)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Date de fin", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Chapô", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Contenu", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Adresse de l'événement", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fichier n°1 (programme)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "En savoir plus (lien externe)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Type de la tête de gondole", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Visuel statique", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Image du slider n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Image du slider n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Image du slider n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Projet du slider n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Projet du slider n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Projet du slider n°3", 'UTF-16', 'UTF-8' );
							break;
						case 'photo':
							$header   = array( mb_convert_encoding( "Identifiant (pour import)", 'UTF-16', 'UTF-8' ) );
							$header[] = mb_convert_encoding( "Identifiant (ID)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Titre", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°4", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°5", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°4", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°5", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fonds", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Auteur (de la prise de vue)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Copyright (propriétaire)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Date de la prise de vue", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Descriptif", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Visuel", 'UTF-16', 'UTF-8' );
							break;
						case 'video':
							$header   = array( mb_convert_encoding( "Identifiant (pour import)", 'UTF-16', 'UTF-8' ) );
							$header[] = mb_convert_encoding( "Identifiant (ID)", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Titre du projet", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Rubrique n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°4", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°5", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°6", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°7", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Thème n°8", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°1", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°2", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Echelon géographique n°3", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Fonds", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Auteur ou source", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Date de la vidéo", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Contenu", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Projet rattaché", 'UTF-16', 'UTF-8' );
							$header[] = mb_convert_encoding( "Vidéo Youtube", 'UTF-16', 'UTF-8' );
							break;
						default:
							$header[] = mb_convert_encoding( $dataField["label"], 'UTF-16', 'UTF-8' );
							break;
					}
				}
			}
			// filename for download
			$filename = $bundle . "_" . date( 'Ymd' ) . ".csv";
			header( "Content-Disposition: attachment; filename=\"$filename\"" );
			header( "Content-Type: text/csv" );
			$out = fopen( "php://output", 'w' );
			fputcsv( $out, $header, ';', '"' );
			foreach ( $data as $row ) {
				//array_walk( $row, __NAMESPACE__ . '\cleanData' );
				fputcsv( $out, array_values( $row ), ';', '"' );
			}
			fclose( $out );
			exit;

		} else {
			if ( is_null( $uid ) ) {
				$ids   = \Drupal::entityQuery( 'user' )
				                ->condition( 'roles', 'annuaire' )
				                ->execute();
				$users = User::loadMultiple( $ids );
			} else {
				$users = User::load( $uid );
			}

			foreach ( \Drupal::entityManager()->getFieldDefinitions( "user", "user" ) as $field_name => $field_definition ) {
				if ( ! empty( $field_definition->getTargetBundle() ) ) {
					$bundleFields[ $field_name ]['type']  = $field_definition->getType();
					$bundleFields[ $field_name ]['label'] = $field_definition->getLabel();
					switch ( $field_name ) {
						case "field_diffusion":
							$bundleFields[ $field_name ]['label'] = "Communautés et instances de travail";
							break;
						default:
							break;
					}
				}
			}
			$data = [];
			foreach ( $users as $cNode ) {
				$elem = [];

				foreach ( $bundleFields as $key => $dataField ) {
					if ( ! in_array( $dataField["type"], $ban ) ) {
						$cleElem = mb_convert_encoding( $key, 'UTF-16', 'UTF-8' );
						if ( $dataField["type"] == "entity_reference" ) {
							$datasEntRef = $cNode->get( $key )->getString();
							$datasEntRef = explode( ", ", $datasEntRef );
							$tmpEntRef   = [];

							if ( ! empty( $cNode->get( $key )->getString() ) ) {
								foreach ( $datasEntRef as $tid ) {
									if ( $term = \Drupal\taxonomy\Entity\Term::load( $tid ) ) {
										$tmpEntRef[] = $term->get( 'name' )->value;
									}
								}
							}

							$valueElem = mb_convert_encoding( implode( ", ", $tmpEntRef ), 'UTF-16', 'UTF-8' );
						} else {
							if ( $dataField["type"] == "file" || $dataField["type"] == "image" ) {
								$valueStringified = $cNode->get( $key )->getString();
								$datasEntRef      = explode( ", ", $valueStringified );

								if ( $dataField["type"] == "image" ) {
									$datasEntRef = array_slice( $datasEntRef, 0, 1 );
								}
								$tmpEntRef = [];
								if ( ! empty( $cNode->get( $key )->getString() ) ) {

									foreach ( $datasEntRef as $fid ) {
										if ( is_numeric( $fid ) && (int) $fid > 1 ) {
											if ( $file = \Drupal\file\Entity\File::load( $fid ) ) {
												$tmpEntRef[] = str_replace( "public://", "https://" . $host . "/sites/default/files/", $file->getFileUri() );
											};
											unset( $file );
										}
									}
								}
								$valueElem = mb_convert_encoding( implode( ", ", $tmpEntRef ), 'UTF-16', 'UTF-8' );
							} else {
								$valueElem = mb_convert_encoding( $cNode->get( $key )->getString(), 'UTF-16', 'UTF-8' );
								if ( $dataField["type"] == "field_collection" ) {
									$valueStringified = $cNode->get( $key )->getString();
									$datasEntRef      = explode( ", ", $valueStringified );
									$tmpEntRef        = [];
									if ( ! empty( $cNode->get( $key )->getString() ) ) {

										foreach ( $datasEntRef as $fcid ) {
											if ( is_numeric( $fcid ) ) {
												if ( $fc = \Drupal\field_collection\Entity\FieldCollectionItem::load( $fcid ) ) {
													$tmpEntRef[] = $fc->get( "field_nom" )->getString();
												};
											}
										}
									}
									$valueElem = mb_convert_encoding( implode( ", ", $tmpEntRef ), 'UTF-16', 'UTF-8' );

								}
							}
						}
						$elem[ $cleElem ] = $valueElem;
					}
				}
				$elem["mail"] = $cNode->get( 'mail' )->value;
				$data[]       = $elem;
			}
			$dataF = array();

			$terms = \Drupal::service( 'entity_type.manager' )->getStorage( "taxonomy_term" )->loadTree( 'diffusion', $parent = 0, $max_depth = null, $load_entities = false );

			foreach ( $terms as $term ) {
				$listeDiff[] = $term->name;
			}
			foreach ( $data as $row ) {
				$headerFile = array(
					"field_nom"          => $row[ mb_convert_encoding( "field_nom", 'UTF-16', 'UTF-8' ) ],
					"field_prenom"       => $row[ mb_convert_encoding( "field_prenom", 'UTF-16', 'UTF-8' ) ],
					"field_sigle"        => $row[ mb_convert_encoding( "field_sigle", 'UTF-16', 'UTF-8' ) ],
					"field_structure"    => $row[ mb_convert_encoding( "field_structure", 'UTF-16', 'UTF-8' ) ],
					"field_service"      => $row[ mb_convert_encoding( "field_service", 'UTF-16', 'UTF-8' ) ],
					"field_fonction"     => $row[ mb_convert_encoding( "field_fonction", 'UTF-16', 'UTF-8' ) ],
					"field_echelons"     => $row[ mb_convert_encoding( "field_echelons", 'UTF-16', 'UTF-8' ) ],
					"field_adresse"      => $row[ mb_convert_encoding( "field_adresse", 'UTF-16', 'UTF-8' ) ],
					"field_telephone"    => $row[ mb_convert_encoding( "field_telephone", 'UTF-16', 'UTF-8' ) ],
					"mail"               => $row["mail"],
					"field_email"        => $row[ mb_convert_encoding( "field_email", 'UTF-16', 'UTF-8' ) ],
					"field_email2"       => $row[ mb_convert_encoding( "field_email2", 'UTF-16', 'UTF-8' ) ],
					"field_website"      => $row[ mb_convert_encoding( "field_website", 'UTF-16', 'UTF-8' ) ],
					"field_presentation" => $row[ mb_convert_encoding( "field_presentation", 'UTF-16', 'UTF-8' ) ],
					"field_projets"      => $row[ mb_convert_encoding( "field_projets", 'UTF-16', 'UTF-8' ) ],
					"field_gal"          => $row[ mb_convert_encoding( "field_gal", 'UTF-16', 'UTF-8' ) ],
				);

				foreach ( $listeDiff as $tid => $tname ) {
					if ( strpos(
						     $row[ mb_convert_encoding( "field_diffusion", 'UTF-16', 'UTF-8' ) ],
						     mb_convert_encoding( $tname, 'UTF-16', 'UTF-8' )

					     ) !== false ) {
						$headerFile[ $tid ] = "X";
					} else {
						$headerFile[ $tid ] = "";
					}
				}
				$dataF[] = $headerFile;
			}

			$header = array(
				mb_convert_encoding( "Nom", 'UTF-16', 'UTF-8' ),
				mb_convert_encoding( "Prénom", 'UTF-16', 'UTF-8' ),
				mb_convert_encoding( "Sigle de la structure", 'UTF-16', 'UTF-8' ),
				mb_convert_encoding( "Dénomination structure", 'UTF-16', 'UTF-8' ),
				mb_convert_encoding( "Direction / Service / Unité", 'UTF-16', 'UTF-8' ),
				mb_convert_encoding( "Fonction du représentant", 'UTF-16', 'UTF-8' ),
				mb_convert_encoding( "Niveau géographique", 'UTF-16', 'UTF-8' ),
				mb_convert_encoding( "Adresse postale", 'UTF-16', 'UTF-8' ),
				mb_convert_encoding( "Télephone fixe", 'UTF-16', 'UTF-8' ),
				mb_convert_encoding( "MAIL 1", 'UTF-16', 'UTF-8' ),
				mb_convert_encoding( "MAIL 2 (Mail générique structure)", 'UTF-16', 'UTF-8' ),
				mb_convert_encoding( "AUTRE EMAIL GENERIQUE DE LA STUCTURE", 'UTF-16', 'UTF-8' ),
				mb_convert_encoding( "Site internet de la structure", 'UTF-16', 'UTF-8' ),
				mb_convert_encoding( "Présentation de la structure", 'UTF-16', 'UTF-8' ),
				mb_convert_encoding( "Projets et thèmes des travaux", 'UTF-16', 'UTF-8' ),
				mb_convert_encoding( "Nom du GAL", 'UTF-16', 'UTF-8' ),
			);
			foreach ( $listeDiff as $tid => $tname ) {
				$header[] = mb_convert_encoding( $tname, 'UTF-16', 'UTF-8' );
			}

			// filename for download
			$filename = $bundle . "_" . date( 'Ymd' ) . ".csv";
			header( "Content-Disposition: attachment; filename=\"$filename\"" );
			header( "Content-Type: text/csv" );
			$out = fopen( "php://output", 'w' );
			fputcsv( $out, $header, ';', '"' );
			foreach ( $dataF as $row ) {
				fputcsv( $out, array_values( $row ), ';', '"' );
			}
			fclose( $out );
			exit;
		}


	}

	public function annuaire(
		$uid
	) {
		set_time_limit( 0 );
		$host = \Drupal::request()->getHost();

		$bundle       = "annuaire";
		$ban          = array(
			"builder",
			'feeds_item',
			'boolean',
			'image'
		);
		$banWkey      = array(
			"field_gps",
			"field_fonds"
		);
		$users        = User::loadMultiple( [ $uid ] );
		$bundleFields = [];
		foreach ( \Drupal::entityManager()->getFieldDefinitions( "user", "user" ) as $field_name => $field_definition ) {
			if ( ! empty( $field_definition->getTargetBundle() ) ) {
				$bundleFields[ $field_name ]['type']  = $field_definition->getType();
				$bundleFields[ $field_name ]['label'] = $field_definition->getLabel();
			}
		}
		/*var_dump( $bundleFields );
		die();*/
		$data = [];
		foreach ( $users as $cNode ) {
			$elem = [];
			foreach ( $bundleFields as $key => $dataField ) {
				if ( ! in_array( $dataField["type"], $ban ) ) {
					if ( ! in_array( $key, $banWkey ) ) {
						$cleElem = mb_convert_encoding( $dataField["label"] . " | " . $dataField["type"], 'UTF-16', 'UTF-8' );
						if ( $dataField["type"] == "entity_reference" ) {
							$datasEntRef = $cNode->get( $key )->getString();
							$datasEntRef = explode( ", ", $datasEntRef );
							$tmpEntRef   = [];

							if ( ! empty( $cNode->get( $key )->getString() ) ) {
								foreach ( $datasEntRef as $tid ) {
									if ( $term = \Drupal\taxonomy\Entity\Term::load( $tid ) ) {
										$tmpEntRef[] = $term->get( 'name' )->value;
									}
								}
							}
							$valueElem = mb_convert_encoding( implode( ", ", $tmpEntRef ), 'UTF-16', 'UTF-8' );

						} else {
							if ( $dataField["type"] == "file" || $dataField["type"] == "image" ) {
								$valueStringified = $cNode->get( $key )->getString();
								$datasEntRef      = explode( ", ", $valueStringified );

								if ( $dataField["type"] == "image" ) {
									$datasEntRef = array_slice( $datasEntRef, 0, 1 );
								}
								$tmpEntRef = [];
								if ( ! empty( $cNode->get( $key )->getString() ) ) {

									foreach ( $datasEntRef as $fid ) {
										if ( is_numeric( $fid ) && (int) $fid > 1 ) {
											if ( $file = \Drupal\file\Entity\File::load( $fid ) ) {
												$tmpEntRef[] = str_replace( "public://", "https://" . $host . "/sites/default/files/", $file->getFileUri() );
											};
											unset( $file );
										}
									}
								}
								$valueElem = mb_convert_encoding( implode( ", ", $tmpEntRef ), 'UTF-16', 'UTF-8' );
							} else {
								$valueElem = mb_convert_encoding( $cNode->get( $key )->getString(), 'UTF-16', 'UTF-8' );
								if ( $dataField["type"] == "field_collection" ) {
									$valueStringified = $cNode->get( $key )->getString();
									$datasEntRef      = explode( ", ", $valueStringified );
									$tmpEntRef        = [];
									if ( ! empty( $cNode->get( $key )->getString() ) ) {

										foreach ( $datasEntRef as $fcid ) {
											if ( is_numeric( $fcid ) ) {
												if ( $fc = \Drupal\field_collection\Entity\FieldCollectionItem::load( $fcid ) ) {
													$tmpEntRef[] = $fc->get( "field_nom" )->getString();
												};
											}
										}
									}
									$valueElem = mb_convert_encoding( implode( ", ", $tmpEntRef ), 'UTF-16', 'UTF-8' );

								}
							}
						}
						$elem[ $cleElem ] = $valueElem;
					}
				}
			}
			$data[] = $elem;
		}

		foreach ( $bundleFields as $key => $dataField ) {
			if ( ! in_array( $dataField["type"], $ban ) && ! in_array( $key, $banWkey ) ) {
				$header[] = mb_convert_encoding( $dataField["label"], 'UTF-16', 'UTF-8' );
			}
		}

		// filename for download
		$filename = $bundle . "_" . date( 'Ymd' ) . ".csv";
		header( "Content-Disposition: attachment; filename=\"$filename\"" );
		header( "Content-Type: text/csv" );
		$out = fopen( "php://output", 'w' );
		fputcsv( $out, $header, ';', '"' );
		foreach ( $data as $row ) {
			fputcsv( $out, array_values( $row ), ';', '"' );
		}
		fclose( $out );
		exit;
	}


}
