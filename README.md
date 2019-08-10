

# Excel to JSON Tool [![Build Status](https://travis-ci.com/vsr061/aem-excel-to-json.svg?branch=master)](https://travis-ci.com/vsr061/aem-excel-to-json)

This project is developed to show how to extend Granite UI to create custom menu under AEM Tools.
#### Concepts covered in this project:
* Sling Resource Merger
* Granite UI Shell

## Tested on

 - AEM 6.5
 - AEM 6.4
 - AEM 6.3

## Download and Install

Download the latest [build](https://github.com/vsr061/aem-excel-to-json/releases) and install through AEM package manager

## Modules

The main parts of the project are:

* core: Java bundle containing all core functionality like OSGi services, listeners or schedulers, as well as component-related Java code such as servlets or request filters.
* ui.apps: contains the /apps (and /etc) parts of the project, ie JS&CSS clientlibs, components, templates, runmode specific configs as well as Hobbes-tests

## How to build

To build all the modules run in the project root directory the following command with Maven 3:

    mvn clean install

If you have a running AEM instance you can build and package the whole project and deploy into AEM with  

    mvn clean install -PautoInstallPackage
    
Or to deploy it to a publish instance, run

    mvn clean install -PautoInstallPackagePublish
    
Or alternatively

    mvn clean install -PautoInstallPackage -Daem.port=4503

Or to deploy only the bundle to the author, run

    mvn clean install -PautoInstallBundle

## Maven settings

The project comes with the auto-public repository configured. To setup the repository in your Maven settings, refer to:

    https://helpx.adobe.com/experience-manager/using/maven_arch13.html
