#!/usr/bin/env tclsh
#-*- mode: Tcl; coding: utf-8-unix; fill-column: 80; -*-

# ----------------------------------------------------------------
# AIDE
# ----------------------------------------------------------------

# Le banc d'essai se lance avec la commande: ` tclsh all.tcl `

# Cette commande lance par défaut tous le fichiers d'essai du répertoire.

# Elle récupère les éventuels paramètres de configuration donnés
# en ligne de commande


package require tcltest

::tcltest::configure -testdir \
    [file dirname [file normalize [info script]]]

if {[llength $argv] > 0} {eval ::tcltest::configure $argv}

::tcltest::runAllTests

::tcltest::cleanupTests 
