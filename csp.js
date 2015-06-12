/*
[crudeSP 0.9.2]
Copyright (C) 2015 Gerald Steinwender This program is free software: you can redistribute it and/or modify it under the terms of the GNU 
General Public License as published by the Free Software Foundation, version. This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public 
License for more details. You should have received a copy of the GNU General Public License along with this program. If not, see 
<http://www.gnu.org/licenses/>. 
*/
var crudeSP = crudeSP || {};
var csp = crudeSP;      //alternative name

//namespace: user
crudeSP.user = crudeSP.user || {};

//namespace: utils
var utils = crudeSP.utils || {};

//namespace: commons
var commons = crudeSP.commons || {};

//////////////////////////////////// U S E R - M O D U L E ///////////////////////////////
//globals in namespace user
crudeSP.user.context = SP.ClientContext.get_current();
crudeSP.user.web = crudeSP.user.context.get_web();

//SP-Username       cannot fetch users of foreign sites 
crudeSP.user.getName = function (callBack, callBackFail) {
    var name = "";
    name = crudeSP.user.web.get_currentUser();
    crudeSP.user.context.load(name);
    crudeSP.user.context.executeQueryAsync(userSuccess, noUsername);

    function noUsername(sender, args) {
        callBackFail(args.get_message());
    }

    function userSuccess() {
        callBack(name.get_title());
    }
};
// /SP-Username

//Groupmembership
crudeSP.user.isMemberOf = function (user, group, callBack, callBackFail) {
    var groups = crudeSP.user.web.get_siteGroups();
    var targetGroup = groups.getByName(group);
    var users = targetGroup.get_users();
    crudeSP.user.context.load(targetGroup);
    crudeSP.user.context.load(users);
    crudeSP.user.context.executeQueryAsync(success, fail);

    function success() {
        var userIsInGroup = false;
        var usersEnumerator = users.getEnumerator();
        while (usersEnumerator.moveNext()) {
            var currUser = usersEnumerator.get_current();
            if (currUser.get_title() == user) {
                userIsInGroup = true;
                break;
            }
        }
        callBack(userIsInGroup);
    }

    function fail() {
        callBackFail("Failed to get groupmembership of user '" + user + "' . Error:" + args.get_message());
    }
}
// /Groupmembership
//////////////////////////////////// /U S E R - M O D U L E ///////////////////////////////


/////////////////////////////////// OutputCaml //////////////////////////////////
crudeSP.returnCaml = function (query, selects) {
    var getCaml = new crudeSP.Operation({
        select: selects,
        where: query,
        list: "none"
    });

    return utils.camlBuilder();
}
/////////////////////////////////// /OutputCaml //////////////////////////////////


/////////////////////////////// O P E R A T I O N //////////////////////////////
crudeSP.Operation = function (definition) {

    //input-properties
    this.list = definition.list;
    this.site = definition.site;                //string or undefined = current
    this.query = definition.where;              //string incl. OrderBy or int (getItemById)
    this.filter = definition.filter;            //object=> {name: string; value: string}  
    this.caml_select = definition.select;       //string -> separation through comma  
    this.orderBy = definition.orderBy;          //string separation through comma              
    this.debugMode = definition.debug || false; //debug consumes runtime, disabled by default
    this.take = definition.take;
    this.set = definition.set;
    this.values = definition.values;
    this.customItemPermissions = true;
    this.roles = definition.roles;      

    //in development  FILEUPLOAD  started on 29.5 14h
    this.file = definition.file;
    this.filename = definition.filename;
    // /in development


    //privates
    var that = this;
    var viewFlds = [];              //this array needs to be globalized because it has to be shared between caml-builder and readItems
    var selectAs = [];              //also needs to be shared between caml-builder and readItems, contains aliases for columnnames
    var caml;                       //global container for computed caml query
    var primaryList;                //primary List
    var listColumns = [];           //queried columnnames of the primary list (debugging: check if field is present)
    var site;                       //global site    \
    var web;                        //global web      ==> to be shared between operations
    var context;                    //global context /
    var affectedRows = 0;           //in updateItems and deleteItems

    var inlineProps = {};           //to be shared between caml builder and read list, constants to written in the result-object
    var caml_selectFields = [];
    var fieldArray = [];

    //regex-constants
    var allowedTabCol = /[\u00c4\u00e4\u00d6\u00f6\u00dc\u00fc\u00df\w\+#!*\/%\?\-\_]+/;        //=> instead of \w+
    var wordDotWord = new RegExp(allowedTabCol.source + "\\." + allowedTabCol.source);          //=> instead of \w+\.\w+
    var wordSEPWord = new RegExp(allowedTabCol.source + "\\_SEP\\_" + allowedTabCol.source);    //=> Seperator für Aliases _SEP_
    var wordSpaceWord = new RegExp(allowedTabCol.source + "\\s+" + allowedTabCol.source);       //=> instead of \w+\s+\w+
    var globDisallow = /[\'\"\s\(\)\[\]\{\}\=\.]+/;                                             //=> currently not allowed: ' " space [ ] ( ) { } = .
    var inlineStatementRgx = new RegExp("\\(" + allowedTabCol.source + "\\s+as\\s+" + allowedTabCol.source + "\\)", "i");             

    //ERROR HANDLING
    //main exception function
    function spqException(type, message, e) {
        this.type = "[crudeSP] " + type;
        this.message = message;
        console.log(this.type + " - " + this.message);
    }

    if (that.debugMode) {
        if (typeof that.list === "undefined") {
            throw new spqException("MAIN - missing parameters", "property 'list' not defined (n. 001)");
        }
    }

    function findObjErrors() {
        if (primaryList.match(wordDotWord)) {
            throw new spqException("MAIN - invalild token", "first list in property 'list' cannot have a table-alias (n. 012)");
        }
        if (that.list.match("join") && typeof that.query === "number") {
            throw new spqException("MAIN - invalid token", "cannot perform a join when query is set to ID (n. 013)");
        }
    }
    // /ERROR HANDLING

    //primary list
    var splitOp = /\s(left[-\s_]?)?join\s/gi;
    primaryList = that.list.split(splitOp)[0];

    if (typeof that.filter !== "undefined" && typeof that.query === "number") {
        that.query = "id = " + that.query;
        //when using bdc-filters, a caml-query has to be passed on to the server, thus getItemById cannot be used
    }
    
    //////////////////////////////////////////// ##### COMMONS ##### /////////////////////////////////////////////////
    //read + update + delete
    commons.GetListAndItems = function () {
        var listItems;
        var list = web.get_lists().getByTitle(primaryList);
        if (typeof that.query === "number") {
            listItems = list.getItemById(that.query);
        } else {
            caml = utils.camlBuilder();
            var qry = new SP.CamlQuery();
            qry.set_viewXml(caml);
            listItems = list.getItems(qry);
        }
        return {
            list: list,
            listItems: listItems
        }
    };

    //read + update + delete
    commons.iteratorFn = function (listItems, execFn) {
        if (typeof listItems.get_count === "undefined") {
            if (typeof listItems.get_fieldValues === "function") {                  
                affectedRows = 1;
            }
        } else {
            affectedRows = listItems.get_count();
        }

        if (affectedRows === 0) {
            console.log("[crudeSP]: MAIN - WARNING: Query returned no results, no rows affected (n. W01)");
        }

        if (affectedRows > 0) {
            if (typeof listItems.getEnumerator === "function") {
                //if listItems count is 1, getEnumerator is undefined
                var enumerateList = listItems.getEnumerator();

                while (enumerateList.moveNext()) {
                    var listItemCurrent = enumerateList.get_current();
                    execFn(listItemCurrent);
                }

            } else {
                execFn(listItems);
            }
        }
    };

    commons.replaceCommas = function (input) {
        var repCommasRegEx = new RegExp("(\\'" + allowedTabCol.source.substring(0, allowedTabCol.source.length - 2) + "\\.\\s,]+\\')");     //=> /(\'[\u00c4\u00e4\u00d6\u00f6\u00dc\u00fc\u00df\w\+#=*\/%\?\-\_\.\s,]+\')/
        if (input.match(repCommasRegEx, "g")) {
            input = input.replace(repCommasRegEx, function (replacement) {
                replacement = replacement.replace(new RegExp(/,/g), "_COMMA_");
                replacement = replacement.replace(new RegExp(/\'/g), "");
                return replacement;
            });
            commons.replaceCommas(input);
        }
        return input;
    }

    commons.stringIsDateFLAWED = function (input) {
        //string + space + number = date; number = date 
        //this method can be problematic as date-recognition is flawed (don't want to use a date-library)
        if (typeof input === "number") {
            return false;
        }
        //long datestring with timezone information is an a object, not a string
        if (typeof input === "object") {
            input = String(input);
        }
        if (!isNaN(new Date(input)) && !input.match(/^[a-zA-Z]+\s\d+/) && !input.match(/^[\d\s]+$/)) {
            return true;
        }
        return false;
    }

    //ERROR HANDLING
    //debug: find field error is in update and create
    commons.findFieldErrors = function (toCheck, failClass, string) {

        var method = "";;
        if (failClass === 2) {
            method = "UpdateItems";

        } else if (failClass === 4) {
            method = "CreateItems";
        }
        if (typeof string === "undefined") {
            string = true;
        }

        if (string) {
            if (toCheck.indexOf("=") === -1) {
                throw new spqException(method + ": invalid token", "no equals operator found in '" + toCheck + "' (n. " + failClass + "11)");
            }
            if (toCheck.match(/=/g).length > 1) {
                throw new spqException(method + ": invalid token", "more then one equals operator found in '" + toCheck + "'. Use ',' to separate columns  (n. " + failClass + "12)");
            }
            toCheck = toCheck.split("=")[0].trim();
            if (toCheck.match(wordDotWord)) {
                throw new spqException(method + ": invalid token", "Table alias found in '" + toCheck + "'. Cannot change columns of non-primary tables: (n. " + failClass + "14)");
            }
            if (toCheck.match(globDisallow)) {
                throw new spqException(method + ": invalid token", "illegal character found in '" + toCheck + "' (n. " + failClass + "13)");
            }
        }
        if (listColumns.indexOf(toCheck) === -1) {
            throw new spqException(method + ": invalid reference: field '" + toCheck + "' not found in List '" + primaryList + "' - fields are case sensitive (n. " + failClass + "21)");
        }

    };

    commons.checkFieldIsExistent = function (input) {
        if (input.match(wordDotWord) || input.match(wordSEPWord)) {
            console.log("[crudeSP]: WARNING: Field '" + input.replace(/_SEP_/, ".") + "' cannot be checked for presence in non-primary list (n. W03)");
            return;
        }

        if (!input.match(wordDotWord) && !input.match(wordSEPWord)) {
            if (listColumns.indexOf(input) === -1) {
                throw new spqException("MAIN - invalid reference", "field '" + input + "' does not exist in list '" + primaryList + "'. Fields are case-sensitive. (n. 021)");
            }
        }
    }

    commons.checkIfUsersGroupsExist = function (source) {
        var users = [];
        var groups = [];
        var rolesRgx = new RegExp(allowedTabCol.source + "[(\\\\)|(\\/)]+");
        var namedSource = "";
        var opName = "";
        if (source === 4) {
            namedSource = "CreateItems";
            opName = "Insert";
        } else if (source === 2) {
            namedSource = "UpdateItems";
            opName = "Update";
        }

        if (typeof that.roles.length === "undefined") {
            var role = that.roles;
            that.roles = [];
            that.roles.push(role);
        }

        utils.restRequest(site + "/_api/web/sitegroups?$select=title", "title").then(function (returnedData) {
            groups = returnedData;
            utils.restRequest(site + "/_api/web/siteusers", "loginname").then(function (moreReturnedData) {
                users = moreReturnedData;
                checkForMembers();
            });
        });

        function checkForMembers() {
            for (var iRoles = 0, iRolesSum = that.roles.length; iRoles < iRolesSum; iRoles++) {
                var offender = that.roles[iRoles].user || that.roles[iRoles].group;
                //ÉRROR HANDLING
                if (that.debugMode) {
                    if (offender.match(wordSpaceWord)) {
                        throw new spqException(namedSource + ": invalid token", "illegal whitespace found in '" + offender + "' use comma to separate user-permissions (n. " + source + "17)");
                    }
                }
                // /ERROR HANDLING
                if (typeof that.roles[iRoles].user !== "undefined") {
                    var userArray = that.roles[iRoles].user.split(",");
                    for (var iUA = 0, iUASum = userArray.length; iUA < iUASum; iUA++) {


                        if (users.indexOf(userArray[iUA].replace(rolesRgx, "").trim()) === -1) { // RGX: /\w+[(\\)|(\/)]+/
                            //throw this message regardless of debug = true to prevent writing data with inherited permissions if a user/group is wrong
                            throw new spqException(namedSource + "invalid reference", "User '" + userArray[iUA] + "' not found. " +opName + " aborted (n. " + source + "422)");
                        }
                    }

                } else if (that.roles[iRoles].group !== "undefined") {
                    var groupArray = that.roles[iRoles].group.split(",");
                    for (var iGA = 0, iGASum = groupArray.length; iGA < iGASum; iGA++) {
                        if (groups.indexOf(groupArray[iGA].replace(rolesRgx, "").trim()) === -1) { // RGX: /\w+[(\\)|(\/)]+/
                            //throw this message regardless of debug = true to prevent writing data with inherited permissions if a user/group is wrong
                            throw new spqException(namedSource + "invalid reference", "User '" + groupArray[iGA] + "' not found. " + opName + " aborted (n. " + source + "422)");
                        }
                    }
                }
            }
        }
    };

    

    //////////////////////////////////////////// ##### /COMMONS ##### /////////////////////////////////////////////////

    ///////////////////////////////// ##### U T I L I T I E S ##### ///////////////////////////////////////
    utils.handleInlineExp = function (index, mode) {
        if (index == 0 || mode === "builder") {
            if (caml_selectFields.length > 0) {
                fieldArray = caml_selectFields;
            } else {
                fieldArray = that.caml_select.split(",");
            }
        }
        if (fieldArray[index].match(inlineStatementRgx)) {
            var match = inlineStatementRgx.exec(fieldArray[index])[0];
            fieldArray.splice(index, 1);
            var parts = match.replace(/\(|\)/g, "").split(/\s+as\s+/i);

            if (!isNaN(parts[0])) {
                parts[0] = parseInt(parts[0]);
            } else if (commons.stringIsDateFLAWED(parts[0])) {
                parts[0] = new Date(parts[0]).toISOString();
            }

            inlineProps[parts[1].trim()] = parts[0];
        }
    };



    utils.encodeNames = function (toEncode) {
        //improvement of foreign code: http://www.n8d.at/blog/encode-and-decode-field-names-from-display-name-to-internal-name/
        //encode special chars 
        var charToEncode = toEncode.split('');
        var encodedString = "";

        for (var i = 0; i < charToEncode.length; i++) {
            var encodedChar = escape(charToEncode[i]);
            if (encodedChar.length > 1) {
                encodedChar = encodedChar.toLowerCase();    //die Zeichenkette, mit der das Sonderzeichen überschreiben wird muss lowerCase sein.
            }

            if (encodedChar.length == 3) {
                encodedString += encodedChar.replace("%", "_x00") + "_";
            }
            else if (encodedChar.length == 5) {
                encodedString += encodedChar.replace("%u", "_x") + "_";
            }
            else {
                encodedString += encodedChar;
            }
        }
        return encodedString;
    }

    ////////////////////// get siteusers and groups ////////////////////////
    utils.restRequest = function (endpoint, resource) {
        var deferred = new $.Deferred();
        var outputContainer = [];
        var request = new XMLHttpRequest();
        request.open("GET", endpoint, true);
        request.setRequestHeader("Accept", "application/json; odata=verbose");

        request.onreadystatechange = function () {
            if (request.readyState === 4) {
                var data = JSON.parse(request.responseText);
                for (var iRes = 0, iSum = data.d.results.length; iRes < iSum; iRes++) {
                    if (resource === "title") {
                        outputContainer.push(data.d.results[iRes].Title);
                    } else if (resource === "loginname") {
                        outputContainer.push(data.d.results[iRes].LoginName.replace(/[\w\:\.\|\#]+\\/, "")); //auth-tag + domain + \ löschen
                    }
                }
                deferred.resolve(outputContainer);
            }
        };
        request.send(null);
        return deferred.promise();
    };

    ////////////////////// /get siteusers and groups ////////////////////////

    ///////////////////// SP-Location /////////////////////
    utils.getContextWeb = function () {

        var deferred = new $.Deferred();
        //ERROR HANDLING
        if (that.debugMode) {
            findObjErrors();
            if (typeof SP === "undefined") {
                throw new spqException("missing prerequisites", "Sharepoint client-side libraries not loaded properly (n. 091)");
            }
        }
        // /ERROR HANDLING

        if (typeof that.site === "undefined") {
            context = SP.ClientContext.get_current();
            site = window.location.href.substring(0, window.location.href.indexOf("SitePages/Home.aspx")); ////site = window.location.protocol + "//" + window.location.host;
        } else {
            context = new SP.ClientContext(that.site);
            site = that.site;
        }
        web = context.get_web();

        if (that.debugMode) {

            utils.restRequest(site + "/_api/web/lists/getByTitle('" + primaryList + "')/fields", "title").then(gotListColumnNames);
        } else {
            gotListColumnNames();
        }

        function gotListColumnNames(data) {
            deferred.resolve(data);
        }

        return deferred.promise();

    };
    ///////////////////// /SP-Location /////////////////////

    ///////////////////// ItemLevel-Perms /////////////////////


    utils.deleteRoles = function (listItem) {
        var deferred = new $.Deferred();

        var perms = listItem.get_roleAssignments();
        listItem.breakRoleInheritance(true);
        context.load(perms);
        context.executeQueryAsync(function () {
            var assignments = [];
            var enumeratePerms = perms.getEnumerator();
            while (enumeratePerms.moveNext()) {
                assignments.push(enumeratePerms.get_current());
            }

            for (var iAss = 0, iAssSum = assignments.length; iAss < iAssSum; iAss++) {
                assignments[iAss].deleteObject();
            }

            context.executeQueryAsync(function () {
                deferred.resolve();
            }, function (sender, args) {
                console.log(args.get_message());
            });

        }, function (sender, args) {
            console.log(args.get_message());
        });
        return deferred.promise();
    };


    utils.assignRoles = function (listItemsToPermit, callBackFn) {
        var userWeb = context.get_web();
        var roleTypesInterface = ["read", "write", "contribute"]; //SP also offers: administrator, guest, none, webDesigner
        var roleTypesSP = [SP.RoleType.reader, SP.RoleType.editor, SP.RoleType.contributor];
        var roles = [];

        if (typeof that.roles.length === "undefined") {
            roles.push(that.roles);
        } else {
            roles = that.roles;
        }

        for (var iR = 0, iRSum = roles.length; iR < iRSum; iR++) {
            //ERROR HANDLING
            if (that.debugMode) {
                if (roleTypesInterface.indexOf(roles[iR].type) === -1) {
                    throw new spqException("AssignRoles: invalid token", "role '" + roles[iR] + "'not valid. (use read, write, contribute) (n. 511)");
                }
                if (typeof roles[iR].user !== "undefined" && typeof roles[iR].group !== "undefined") {
                    throw new spqException("AssignRoles: invalid token", "cannot have 'user' and 'group' in the same instance (n. 512)");

                }
                if (typeof roles[iR].user !== "undefined") {
                    if (roles[iR].user.match(wordSpaceWord)) {
                        throw new spqException("AssignRoles: invalid token", "Users not separated with comma at " + roles[iR].user + " (n. 513)");
                    }
                } else if (typeof roles[iR].group !== "undefined") {
                    if (roles[iR].group.match(wordSpaceWord)) {
                        throw new spqException("AssignRoles: invalid token", "Groups not separated with comma at " + roles[iR].group + " (n. 513)");
                    }
                }
            }
            // /ERROR HANDLING

            var userArray;
            var groupArray;
            var user_group;
            var sum;

            if (typeof roles[iR].user !== "undefined") {
                userArray = roles[iR].user.split(",");
                sum = userArray.length;
            } else {
                groupArray = roles[iR].group.split(",");
                sum = groupArray.length;
            }


            for (var iUA = 0; iUA < sum; iUA++) {
                if (typeof userArray !== "undefined") {
                    //ACHTUNG: \t \n \r \f \b in strings are not escaped (problematic when entering domain and user name beginning with char)
                    var escapeCharsRgx = [/\n/, /\r/, /\t/, /\f/];
                    var escapeCharsStr = ["n", "r", "t", "f"];
                    var userToEnsure;
                    for (var iEC = 0; iEC < 6; iEC++) {
                        if (userArray[iUA].match(escapeCharsRgx[iEC])) {
                            userToEnsure = userArray[iUA].replace(escapeCharsRgx[iEC], "\\" + escapeCharsStr[iEC]);
                            break;
                        }
                    }
                    userToEnsure = userToEnsure.replace(/[\/]+/, "\\");
                    user_group = userWeb.ensureUser(userToEnsure);
                } else {
                    user_group = userWeb.get_siteGroups().getByName(groupArray[iUA].trim());
                }

                var roleDefBinding = SP.RoleDefinitionBindingCollection.newObject(context);

                for (var iRT = 0; iRT < 3; iRT++) {
                    if (roles[iR].type === roleTypesInterface[iRT]) {
                        roleDefBinding.add(userWeb.get_roleDefinitions().getByType(roleTypesSP[iRT]));

                        for (var iLI = 0, iLISum = listItemsToPermit.length; iLI < iLISum; iLI++) {
                            var listItem = listItemsToPermit[iLI];

                            listItem.breakRoleInheritance(false);

                            listItem.get_roleAssignments().add(user_group, roleDefBinding); //<= add(user, roleDefBinding)

                        }
                        break;
                    }
                }
            } //users
        } //roles


        context.executeQueryAsync(function () {
            //console.log("item level permissions set");
            callBackFn();
        }, permFail);

        function permFail(sender, args) {
            console.log(args.get_message());
        }
    };


    ///////////////////// /ItemLevel-Perms /////////////////////

    ///////////////////// C A M L   B U I L D E R /////////////////////
    utils.camlBuilder = function () {

        //ERROR HANDLING
        if (that.debugMode) {

            (function checkInputErrors() {
                if (typeof that.caml_select === "undefined" && typeof that.query === "undefined") {
                    throw new cbException("missing parameters", "no query entered (n. 901)");
                }
                
                if (typeof that.query !== "undefined") {
                    if (that.query.match(/\(/g) || that.query.match(/\)/g)) {                           //*do if there is at least one bracket:
                        if (that.query.match(/\(/g) && that.query.match(/\)/g)) {                       //**if there are opening and closing brackets:
                            if (that.query.match(/\(/g).length !== that.query.match(/\)/g).length) {    //if their count is not even => Ex
                                throw new cbException("invalid token", "discrepancy between opening and closing brackets (n. 911)");
                            }
                        } else {                                                                        //***there is only one sort bracket => Ex
                            throw new cbException("invalid token", "discrepancy between opening and closing brackets (n. 911)");
                        }
                        var stringByBracket = that.query.split("(");
                        for (var iSBB = 1, iSBBSum = stringByBracket.length; iSBB < iSBBSum; iSBB++) {  //start at index = 1, because index 0 is the string to the first bracket
                            if (stringByBracket[iSBB].indexOf("and") === -1 && stringByBracket[iSBB].indexOf("or") === -1) {
                                throw new cbException("invalid token", "logical operand missing at: " + stringByBracket[iSBB] + " (n. 912)");
                            }
                        }
                    }
                }
                if (typeof that.take !== "undefined") {
                    if (that.take % 1 !== 0 || that.take < 0) {
                        throw new cbException("invalid token", "Value for take must be a positive integer (n. 913)");
                    }
                }
            })();

        }


        function findIllegalTokens(input, mode) {
            if (that.debugMode) {
                if (mode === 1) {
                    if (input.match(/[\.]{2}/g) || input.match(/[\_]{2}/g)) { // . and _ cannot occurr twice
                        throw new cbException("invalid token", "too much '.' or '_' found in: " + input + " (n. 915)");
                    }
                } else if (mode > 1) {
                    if (input.match(globDisallow)) {
                        if (mode === 2) {
                            throw new cbException("invalid token", "illegal character found in alias: " + input + " (n. 9112)");
                        } else if (mode === 3) {
                            if (input.match(wordDotWord))

                            throw new cbException("invalid token", "illegal character found in field: " + input + " (n. 9113)");
                        }
                    }
                }
            }
        }

        //exception-function-CAML-Builder
        function cbException(type, message) {
            this.type = "[crudeSP] CAML-BUILDER - " + type;
            this.message = message;
            console.log(this.type + " - " + this.message);
        }

        // /ERROR HANDLING

        var validCaml = ""; //final caml output

        var operands = {
            straightOperands: ["<>", "=", "!=", "<", "<=", ">", ">=", "IS NULL", "IS NOT NULL", "CONTAINS", "BEGINS WITH", "IN", "and", "or"],              //and + or need a regex to prevent matching within strings
            straightOperandsRgx: [/<>/, /=/, /!=/, /</, /<=/, />/, />=/, /IS\s+NULL/, /IS\s+NOT\s+NULL/, /\s+CONTAINS\s+/, /BEGINS\s+WITH/, /\s+IN\s+/, /(\s+and\s+)|(^and[\s+]?$)/, /(\s+or\s+)|(^or[\s+]?$)/],
            camlOperands: ["<Neq>", "<Eq>", "<Neq>", "<Lt>", "<Leq>", "<Gt>", "<Geq>", "<IsNull>", "<IsNotNull>", "<Contains>", "<BeginsWith>", "<In>", "<And>", "<Or>"],
            camlOperandsClosed: ["</Neq>", "</Eq>", "</Neq>", "</Lt>", "</Leq>", "</Gt>", "</Geq>", "</IsNull>", "</IsNotNull>", "</Contains>", "</BeginsWith>", "</In>", "</And>", "</Or>"]
        };

        var viewFields = [];        //if there is no select property set, take fields within this array as viewfields
        var orderByFromString = []; //debug
        var aliases = [];           //debug
        var aliasesWhere = [];      //debug
        var aliasesOrderBy = [];    //debug
        var aliasesSelect = [];     //debug

        //INIT
        (function initialze() {
            headerAndFilter();

            //uniformize keywords (case sensivity)
            if (typeof that.query !== "undefined") {
                that.query = that.query.replace(/\s+and\s+/gi, " and ");
                that.query = that.query.replace(/\s+or\s+/gi, " or ");
                that.query = that.query.replace(/is\snull/gi, "IS NULL");
                that.query = that.query.replace(/is\snot\snull/gi, "IS NOT NULL");
                that.query = that.query.replace(/\s+contains\s+/gi, " CONTAINS ");
                that.query = that.query.replace(/begins\swith/gi, "BEGINS WITH");
                that.query = that.query.replace(/\s+in\s+/gi, " IN ");

                convertWhereString(that.query);
            } else {
                convertWhereString("");
            }

        })();
        // /INIT

        function headerAndFilter() {
            validCaml = "<View>";

            //25.2  top xx rows
            if (typeof that.take !== "undefined") {
                validCaml += "<RowLimit Paged='False'>" + that.take + "</RowLimit>";
            }
            // /25.2
           
            if (typeof that.filter !== "undefined") {
                var method = that.filter.readOperationName || "ReadList";
                validCaml += "<Method Name='" + method + "'>";           //SP-Designer calls it Read List whereas visual studio calls it ReadList
            }
            if (typeof that.filter !== "undefined") {
                validCaml += "<Filter Name='" + that.filter.name + "' Value='" + that.filter.value + "' />";
            }
            if (typeof that.filter !== "undefined") {
                validCaml += "</Method>";
            }
            validCaml += "<Query>";
            if (typeof that.query !== "undefined") {
                validCaml += "<Where>";
            }
        }


        function convertWhereString(qryString) {
            var output = [];

            (function slicer(queryString) {

                var n_and = (queryString.match(/\s+and\s+|^and\s+/g) || []).length;
                var n_or = (queryString.match(/\s+or\s+/g) || []).length;            //new match: /\s+or\s+|^or\s+/g
                var n_brckOp = (queryString.match(/\(/g) || []).length;
                var n_brckCls = (queryString.match(/\)/g) || []).length;

                //breaking the recursion
                if (n_and == 0 && n_or == 0 && n_brckOp == 0 && n_brckCls == 0) {
                    output.push(queryString);
                    return;
                }

                //create an array of operands, sorted by their occurrance 
                var position = [];
                (function positioning() {
                    for (var iPos = 0; iPos < 4; iPos++) {
                        var operand;
                        if (iPos === 0) {
                            operand = " and ";
                        } else if (iPos === 1) {
                            operand = " or ";
                        } else if (iPos === 2) {
                            operand = "(";
                        } else if (iPos === 3) {
                            operand = ")";
                        }

                        //when char not found, indexOf = -1 
                        var index = 10000;
                        if (queryString.indexOf(operand) !== -1) {
                            index = queryString.indexOf(operand);
                        }

                        position[iPos] = {
                            index: index,
                            sliceby: operand,
                            slicebyLen: operand.length
                        }
                    }

                    position.sort(function (a, b) {
                        return a.index - b.index;
                    });

                })();

                var condition = queryString.substr(0, queryString.indexOf(position[0].sliceby)).trim();

                if (condition !== "") {
                    output.push(condition);
                }
                output.push(position[0].sliceby.trim());
                slicer(queryString.substr(queryString.indexOf(position[0].sliceby) + position[0].slicebyLen).trim()); //REKURSION bis kein 'and', 'or' oder '(' bzw. ')' im string ist.
            })(qryString);
            processWhereArray(output); //vv
        }

        function processWhereArray(qryArray) {
            var camlWh = [];
            for (var i = 0, iSum = qryArray.length; i < iSum; i++) { //improving performence by setting array length within for-loop-init

                if (qryArray[i] !== "and" && qryArray[i] !== "or" && qryArray[i] !== "(" && qryArray[i] !== ")" && qryArray[i].trim() !== "") {
                    var slice = qryArray[i];
                    var operand = translateOperand(slice);
                    var sliceLeftArr = slice.split(operand.normalOps)[0].trim().split("."); //value before operand
                    //ERROR HANDLING
                    if (that.debugMode) {
                        if (slice.match(/\'/)) {
                            if (slice.match(/\'/g).length % 2 !== 0) {
                                throw new cbException("invalid token", "no closing ' found at " + slice + " (n. 9115)");
                            }
                        }
                        checkForMissingOperands(slice);
                        var field = slice.split(operand.normalOps)[0].trim();
                        findIllegalTokens(field, 1);
                        commons.checkFieldIsExistent(field);


                    }
                    // /ERROR HANDLING

                    //strings can be enclosed with ' (to escape special chars within strings like a comma)
                    slice = slice.replace(new RegExp(/\'/g), "");
                    var sliceLeft = "";
                    debugger;
                    if (sliceLeftArr.length === 2) {
                        aliasesWhere.push(sliceLeftArr[0]); //debug
                        sliceLeft = sliceLeftArr[0] + "_SEP_" + sliceLeftArr[1]; //replace alias with _SEP_ (separator)]
                    } else {
                        sliceLeft = sliceLeftArr.join("");
                    }

                    var sliceRight = slice.split(operand.normalOps)[1].trim(); //value after operand
                    if (viewFields.indexOf(sliceLeft) === -1) {
                        viewFields.push(sliceLeft); //if no select columns are set, use this array for viewfields
                    }

                    //I. use this pattern when using IN
                    if (operand.normalOps === "IN") {
                        //ERROR HANDLING
                        if (that.debugMode) {
                            (function findInEror() {
                                if (sliceRight.match(/\d+\s+\d+/)) {    // number(s) space number(s)
                                    throw new cbException("invalid token", "IN: numbers not separated by commas (n. 913)");
                                }
                            })();
                        }
                        // /ERROR HANDLING
                        camlWh.push(operand.camlOps);
                        camlWh.push("<FieldRef Name = '" + sliceLeft + "' />");      //=> SPENCODE
                        camlWh.push("<Values>");

                        //2.4.2015 IN for strings
                        var inIntegers = sliceRight.split(",");
                        for (var iI = 0, iISum = inIntegers.length; iI < iISum; iI++) {
                            var typIn = checkType(inIntegers[iI].trim());
                            var counter = "<Value Type='" + typIn.typ + "'>";
                            camlWh.push(counter + inIntegers[iI].trim() + "</Value>");
                        }
                        camlWh.push("</Values>");
                        // /2.4.2015

                        camlWh.push(operand.camlOpsClosed);
                    } else { //II. normal pattern without IN
                        camlWh.push(operand.camlOps);
                        camlWh.push("<FieldRef Name = '" + utils.encodeNames(sliceLeft) + "' />");      //=> SPENCODE
                        var typ = checkType(sliceRight);
                        camlWh.push("<Value Type='" + typ.typ + "'>" + typ.value + "</Value>");
                        camlWh.push(operand.camlOpsClosed);
                    }
                } else if (qryArray[i] === "and" || qryArray[i] === "or") {
                    camlWh.push(translateOperand(qryArray[i]).camlOps);
                } else if (qryArray[i] === "(" || qryArray[i] === ")") {
                    camlWh.push(qryArray[i]);
                }
            }

            sortArrayToCaml(camlWh);

            //ERROR HANDLING
            function checkForMissingOperands(string) {
                var operandFound = false;
                for (var iOperand = 0, iOperandSum = operands.straightOperands.length; iOperand < iOperandSum; iOperand++) {
                    if (string.indexOf(operands.straightOperands[iOperand]) !== -1) {
                        operandFound = true;
                    }
                }
                if (!operandFound) {
                    throw new cbException("invalid token", "no comparing operand in: " + string + " (n. 914)");
                }
            }
            // /ERROR HANDLING

            function translateOperand(inpu) {
                var toReturn = {};
                for (var iSub = 0, iSumSub = operands.straightOperands.length; iSub < iSumSub; iSub++) {

                    if (inpu.match(operands.straightOperandsRgx[iSub])) {
                        toReturn.camlOps = operands.camlOperands[iSub];
                        toReturn.normalOps = operands.straightOperands[iSub];
                        toReturn.camlOpsClosed = operands.camlOperandsClosed[iSub];
                    }
                }
                return toReturn; //returns an array 0 caml operand; 1 original operand 
            };

            function checkType(ipt) {
                var toReturn = {};
                toReturn.value = ipt;
                if (!isNaN(ipt)) {
                    if (ipt % 1 == 0) {
                        toReturn.typ = "Integer";
                    } else toReturn.typ = "Number";
                } else if (commons.stringIsDateFLAWED(ipt)) { 
                    toReturn.typ = "DateTime";
                    toReturn.value = new Date(ipt).toISOString();
                } else toReturn.typ = "Text";

                return toReturn; //returns an objcet (typ, value); if value is DateTime return the ISO-string
            }

        }

        //ORDER BY
        function orderByFn() {
            var orderBy = "";
            if (typeof that.orderBy !== "undefined") {
                orderBy = "<OrderBy>";
                var splittedCaml_orderBy = that.orderBy.split(",");
                for (var iOb = 0, iObSum = splittedCaml_orderBy.length; iOb < iObSum; iOb++) {
                    var itemAndDirection = splittedCaml_orderBy[iOb].trim();
                    //ERROR HANDLING
                    if (that.debugMode) {
                        (function () {
                            var checkObFld = splittedCaml_orderBy[iOb].replace(/\s+asc|\s+desc/gi, ""); //für das orderBy caml_selectFields[iCS].replace(/\s(desc|asc)/gi, "");
                            if (checkObFld.match(wordSpaceWord)) {
                                throw new cbException("invalid token", "orderBy fields need to be divided by comma (n. 9111)");
                            }
                        })();
                    }
                    // /ERROR HANDLING
                    var item = itemAndDirection.split(/\s+/)[0].replace(/\./g, "_SEP_");
                    if (item.match(/_SEP_/)) {
                        aliasesOrderBy.push(item.split("_SEP_")[0]); //debug  
                    }
                    var direction2Caml = "TRUE";
                    if (itemAndDirection.match(/\s+desc/i)) {
                        direction2Caml = "FALSE";
                    }
                    orderByFromString.push(item); //-->needed later for ERROR HANDLING
                    orderBy += "<FieldRef Name='" + utils.encodeNames(item) + "' Ascending='" + direction2Caml + "' />";        //=> SPENCODE
                }
                orderBy += "</OrderBy>";
            }
            orderBy += "</Query>";
            return orderBy;
        }
        // /ORDER BY

        //CAML JOIN             creates XML-Properties <Joins> and <ProjectedFields>
        function joins() {
            var joinCaml = "";
            var projectedFlds = "";
            if (typeof that.list !== "undefined") {

                if (that.list.match(/\s+(left[-\s_]?)?join\s+/gi)) {
                    joinCaml = (function () {
                        var joinsCaml = "<Joins>";
                        that.list = that.list.replace(/\s+join\s+/gi, " join ");
                        that.list = that.list.replace(/\s+left[-\s+_]?join\s+/gi, " left-join ");
                        that.list = that.list.replace(/\s+with\s+/gi, " with ");

                        var splittedFrom = that.list.split(",");
                        var parentList = "";
                        for (var iFrom = 0, iFromSum = splittedFrom.length; iFrom < iFromSum; iFrom++) {
                            var splitBy = /join/;
                            var type = "INNER";
                            if (splittedFrom[iFrom].match(/left[-\s_]?join/i)) {
                                splitBy = /left[-\s_]?join/i;
                                type = "LEFT";
                            }
                            var listSrc = splittedFrom[iFrom].split(splitBy)[0].trim();
                            var listJoin = splittedFrom[iFrom].split(splitBy)[1].trim();
                            var aliasNforeignKey = listJoin.split("with");
                            debugger;
                            var currentAlias = aliasNforeignKey[0].trim();
                           
                            //ERROR HANDLING 
                            if (that.debugMode) {
                                (function checkJoin() {
                                    //right-join is not supported in CAML
                                    if (that.list.match(/right[-\s+_]?join/gi)) {
                                        throw new cbException("not supported", "right join is not supported by the system (n. 931)");
                                    }
                                    if (splittedFrom[iFrom].match(/left[-\s+_]?join[\w\s]+?((left[-\s+_]?)?join)/i)) {
                                        throw new cbException("invalid token", "join-operations need to be devided with a comma (n. 917)");
                                    }
                                })();
                                findIllegalTokens(currentAlias, 2);
                            }
                            // /ERROR HANDLING
                            aliases.push(currentAlias);
                            var foreignKey = aliasNforeignKey[1].trim().replace(/.*\./, "");

                            if (iFrom === 0) {
                                parentList = listSrc;
                            }
                            var printListSrc = "";

                            if (listSrc !== parentList && listSrc !== "") {
                                printListSrc = "List='" + listSrc + "'";                                       
                            }
                            joinsCaml += "<Join Type='" + type + "' ListAlias='" + currentAlias + "' >";
                            joinsCaml += "<Eq>";
                            joinsCaml += "<FieldRef " + printListSrc + " Name='" + utils.encodeNames(foreignKey) + "' RefType='Id' />";     //=>SPENCODE
                            joinsCaml += "<FieldRef List='" + currentAlias + "' Name='ID' />";
                            joinsCaml += "</Eq>";
                            joinsCaml += "</Join>";
                        }

                        joinsCaml += "</Joins>";
                        return joinsCaml;

                    })();



                    projectedFlds = (function () {
                        if (typeof that.caml_select !== "undefined") {
                            caml_selectFields = that.caml_select.split(",");
                        } else {
                            caml_selectFields = viewFields;
                        }
                        var projFieldsCaml = "<ProjectedFields>";
                        for (var aI = 0, aISum = aliases.length; aI < aISum; aI++) {
                            for (var iSF = 0, iSFSum = caml_selectFields.length; iSF < iSFSum; iSF++) {
                                if (caml_selectFields[iSF].indexOf(".") !== -1 || caml_selectFields[iSF].indexOf("_SEP_") !== -1) {
                                    var splitBy = ".";
                                    if (caml_selectFields[iSF].indexOf("_SEP_") !== -1) {
                                        splitBy = "_SEP_";
                                    }
                                    var ali = caml_selectFields[iSF].split(splitBy)[0].trim();
                                    var column = caml_selectFields[iSF].split(splitBy)[1].trim().replace(/\s+as\s+/gi, " as ");

                                    if (column.match(/\s+as\s+[^.]*/i)) {
                                        column = column.split(/\s+as\s+/i)[0]; //split(" as ")
                                    }
                                    if (ali === aliases[aI]) {
                                        projFieldsCaml += "<Field ShowField='" + utils.encodeNames(column) + "' Type='Lookup' Name='" + ali + "_SEP_" + utils.encodeNames(column) + "' List='" + ali + "' />"; //=>SPENCODE      alias-Spalte mit _ trennen, das liest auch caml
                                        var theAs = "";
                                        if (caml_selectFields[iSF].match(/\s+as\s+[^.]*/i)) {
                                            theAs = " as " + caml_selectFields[iSF].split(/\s+as\s+/i)[1];       //split(" as ")
                                        }
                                        caml_selectFields.splice(iSF, 1, ali + "_SEP_" + column + theAs);
                                    }
                                }
                            }
                        }
                        projFieldsCaml += "</ProjectedFields>";
                        return projFieldsCaml;
                    })();
                }
            }
            return joinCaml + projectedFlds;
        }

        // /CAML JOIN

        //VIEWFIELDS
        function viewFieldsFn() {
            var viewFieldsCaml = "<ViewFields>";
            if (typeof that.caml_select === "undefined") {
                //DEFAULT-BEHAVIOR: no viewfields set, take the fields used in where
                selectAs = viewFields.slice(0);
                generateVFCaml(viewFields);
            } else {
                var viewFieldsFromString = [];
                if (caml_selectFields.length === 0) {
                    caml_selectFields = that.caml_select.split(",");
                }
                for (var iCS = 0, iCSSum = caml_selectFields.length; iCS < iCSSum; iCS++) {
                    var baseField = caml_selectFields[iCS].trim();
                    //30.3 -inline Statements
                    if (caml_selectFields[iCS].match(inlineStatementRgx)) {
                        utils.handleInlineExp(iCS, "builder");
                        // /30.3 - inline Statements
                    } else {
                        baseField = baseField.replace(/\s+as\s+/gi, " as ");
                        //ERROR HANDLING
                        if (that.debugMode) {
                            (function () {
                                var checkSelFld = caml_selectFields[iCS].replace(new RegExp("\\s+as\\s+" + allowedTabCol.source, "gi"), "").trim();

                                //30.3
                                if (caml_selectFields[iCS].match(/\(/g) || caml_selectFields[iCS].match(/\)/g)) {
                                    if (caml_selectFields[iCS].match(/\(/g) && caml_selectFields[iCS].match(/\)/g)) {
                                        if (caml_selectFields[iCS].match(/\(/g).length !== caml_selectFields[iCS].match(/\)/g).length) {
                                            throw new cbException("invalid token", "discrepancy between opening and closing brackets found at '" + caml_selectFields[iCS].trim() + "' (n. 9116)");
                                        }
                                    } else {
                                        throw new cbException("invalid token", "discrepancy between opening and closing brackets found at '" + caml_selectFields[iCS].trim() + "' (n. 9116)");
                                    }

                                    if (!caml_selectFields[iCS].match(new RegExp(allowedTabCol.source + "\\s+as\\s+", "i"))) {
                                        throw new cbException("invalid token", "no 'as' found in inline expression '" + caml_selectFields[iCS].trim() + "' (n. 9117)");
                                    }
                                }

                                commons.checkFieldIsExistent(checkSelFld);

                                // /30.3
                                if (checkSelFld.match(wordSpaceWord)) {
                                    throw new cbException("invalid token", "select fields need to be divided by comma (n. 918)");
                                }
                                if (caml_selectFields[iCS].match(/^\s+?as/gi)) {
                                    throw new cbException("invalid token", "fieldname cannot be 'as' (n. 919)");
                                }
                                if (that.list.match(/\s(left[-\s+_]?)?join\s+/gi) && that.caml_select.match(/_SEP_/g)) {
                                    throw new cbException("invalid token", "field: '" + caml_selectFields[iCS] + "' cannot be contain '_SEP_' when using join (n. 9110)");
                                }
                            })();
                        }
                        // /ERROR HANDLING

                        //select as
                        if (baseField.match(/\sas\s[^.]*/)) {
                            var splittedBaseField = baseField.split(/\s+as\s+/i); //split(" as ")
                            selectAs.push(splittedBaseField[1].trim());
                            viewFieldsFromString.push(splittedBaseField[0].trim());
                        } else {
                            selectAs.push(baseField);
                            viewFieldsFromString.push(baseField); 
                        }
                        // /select as
                    }
                }
                generateVFCaml(viewFieldsFromString);
            }

            function generateVFCaml(fields) {
                for (var iVF = 0, iVFSum = fields.length; iVF < iVFSum; iVF++) {
                    fields[iVF] = fields[iVF];       //=>SPENCODE
                    debugger;
                    //aliases checken hier!
                    if (fields[iVF].match(wordDotWord) || fields[iVF].match(wordSEPWord)) {
                        aliasesSelect.push(fields[iVF].split(/\.|\_SEP\_/)[0]);
                    }

                    //ERROR HANDLING
                    if (that.debugMode) {
                        (function findAliasErrors() {
                            debugger;
                            for (var iAl = 0, iAlSum = aliases.length; iAl < iAlSum; iAl++) {
                                if (aliases[iAl].match(globDisallow)) {
                                    throw new cbException("invalid token", "illegal character found in alias: " + aliases[iAl] + "  (n. 9112)");
                                }
                            }
                            for (var iAOb = 0, iAObSum = aliasesOrderBy.length; iAOb < iAObSum; iAOb++) {
                                if (aliases.indexOf(aliasesOrderBy[iAOb]) === -1) {
                                    throw new cbException("invalid reference", "alias: " + aliasesOrderBy[iAOb] + " in orderBy not found in list (n. 921)");
                                }
                            }
                            for (var iAW = 0, iAWSum = aliasesWhere.length; iAW < iAWSum; iAW++) {
                                if (aliases.indexOf(aliasesWhere[iAW]) === -1) {
                                    throw new cbException("invalid reference", "alias: " + aliasesWhere[iAW] + " in where not found in list (n. 922)");
                                }
                            }
                            for (var iSel = 0, iSelSum = aliasesSelect.length; iSel < iSelSum; iSel++) {
                                if (aliases.indexOf(aliasesSelect[iSel]) === -1) {
                                    throw new cbException("invalid reference", "alias: " + aliasesSelect[iSel] + " in select not found in list (n. 924)");
                                }
                            }

                        })();
                    }
                    // /ERROR HANDLING


                    // /aliases checken


                    findIllegalTokens(fields[iVF], 3);
                    viewFlds.push(fields[iVF]);     //ViewFields for readItems method (global)
                    viewFieldsCaml += "<FieldRef Name ='" + utils.encodeNames(fields[iVF]) + "' />";
                }
            }

            //ERROR HANDLING
            if (that.debugMode) {
                (function checkOrderByFlds() {
                    for (var iO = 0, iOSum = orderByFromString.length; iO < iOSum; iO++) {

                        if (viewFlds.indexOf(orderByFromString[iO]) === -1) {
                            throw new cbException("invalid reference", "orderBy: field '" + orderByFromString[iO].replace(/_SEP_/, ".") + "' not found in selected fields (n. 923)");
                        }
                    }
                })();
            }
            // /ERROR HANDLING

            viewFieldsCaml += "</ViewFields>";
            return viewFieldsCaml;
        }

        // /VIEWFIELDS

        function sortArrayToCaml(processedArray) {
            var lastOperand;
            var operandChain = [];
            var bracketChain = [];
            var lastPos = 0;
            var lastOpenBracket = 0;


            (function iterateArray(startPos) {
                var hasAndOr = false;
                for (var i = startPos, iSum = processedArray.length; i < iSum; i++) {
                    var setPos = 0;
                    if (lastPos !== 0) {
                        setPos = lastPos + 1;
                    }

                    if (processedArray[i] === "<And>" || processedArray[i] === "<Or>") {
                        hasAndOr = true;
                        lastPos = i;
                        lastOperand = processedArray[i];
                        operandChain.push(operands.camlOperandsClosed[operands.camlOperands.indexOf(processedArray[i])]);
                        processedArray.splice(i, 1);
                        processedArray.splice(setPos, 0, lastOperand);
                        break;
                    }
                    if (processedArray[i] === "(") {
                        lastOpenBracket = i;
                        for (var iBrck = i + 1; iBrck < iSum; iBrck++) {
                            //replace bracket with next and + or
                            if (processedArray[iBrck] === "<And>" || processedArray[iBrck] === "<Or>") {
                                bracketChain.push(operands.camlOperandsClosed[operands.camlOperands.indexOf(processedArray[iBrck])]);
                                processedArray.splice(i, 1, processedArray[iBrck]);
                                processedArray.splice(iBrck, 1);
                                lastPos = iBrck - 1;
                                hasAndOr = true;
                                break;
                                //if there is no and + or within brackets, delete them
                            } else if (processedArray[iBrck] === ")") {
                                processedArray.splice(i, 1);
                                processedArray.splice(iBrck - 1, 1);
                                break;
                            }
                        }
                        break;
                    }

                    if (processedArray[i] === ")") {

                        processedArray.splice(i, 1, bracketChain[bracketChain.length - 1]);
                        bracketChain.splice(-1, 1);
                        if (processedArray[i + 1] === "<And>" || processedArray[i + 1] === "<Or>") {
                            operandChain.push(operands.camlOperandsClosed[operands.camlOperands.indexOf(processedArray[i + 1])]);
                            processedArray.splice(lastOpenBracket, 0, processedArray[i + 1]);
                            processedArray.splice(i + 2, 1); //i+2 after splice in previous line, value is found at i + 2
                            lastPos = i;
                            hasAndOr = true;
                        } else if (processedArray[i + 1] === ")") {
                            hasAndOr = true;
                            lastPos = i;
                            break;
                        }

                        break;
                    }
                }
                if (hasAndOr) {
                    iterateArray(lastPos);
                }
            })(0);

            //closing tags
            (function closeTags() {
                for (var iInv = operandChain.length; iInv >= 0; iInv--) {
                    processedArray.push(operandChain[iInv]);
                }
                if (typeof that.query !== "undefined") {
                    processedArray.push("</Where>");
                }
            })();
            // /closing tags


            //assemble CAML-string
            validCaml += processedArray.join('');
            validCaml += orderByFn();
            validCaml += joins();
            validCaml += viewFieldsFn();
            validCaml += "</View>";
        }

        return validCaml;
    };
    ///////////////////// /C A M L   B U I L D E R /////////////////////

    ///////////////////////////////// ##### /U T I L I T I E S ##### ///////////////////////////////////////

    ///////////////////////////////// ##### CRUD Methods ##### ///////////////////////////////////////

    ///////////////////// ReadItems /////////////////////
    this.readItems = function (callBack, callBackFail) {
        var outputArray = [];
        var listNEntities;
        var list;
        var listItems;
        //ERROR HANDLING
        if (that.debugMode) {
            if (typeof callBack === "undefined") {
                throw new spqException("ReadItems: missing parameters", "no callback function given (n. 101)");
            }
        }
        // /ERROR HANDLING

        utils.getContextWeb().then(function (data) {
            listColumns = data;
            listNEntities = new commons.GetListAndItems();
            list = listNEntities.list;
            listItems = listNEntities.listItems;
            context.load(list);
            context.load(listItems);
            context.executeQueryAsync(readingList, listFail);
        });

        //on success
        function readingList() {
            if (that.debugMode) { //warining in debug-mode
                var warningShown = false;
                var warnAlias = "";
            }

            commons.iteratorFn(listItems, processEntityToObj);

            function processEntityToObj(tableRow) {
                var containerObj = {};
                var properties = [];
                var property;

                if (viewFlds.length === 0) { //caml-builder got bypassed with getItemById 
                    caml_selectFields = that.caml_select.split(",");
                    for (var iCaml = 0, iCamlSum = caml_selectFields.length; iCaml < iCamlSum; iCaml++) {
                        utils.handleInlineExp(iCaml);
                    }
                    viewFlds = caml_selectFields;
                }
                
                for (var iVFlds = 0, iVFldsSum = viewFlds.length; iVFlds < iVFldsSum; iVFlds++) {
                    viewFlds[iVFlds] = viewFlds[iVFlds].trim();
                    if (viewFlds[iVFlds].match(/\s+as\s+[^.]*/i)) {
                        selectAs.push(viewFlds[iVFlds].split(/\s+as\s+/i)[1]);           // split(" as ")
                        viewFlds[iVFlds] = viewFlds[iVFlds].split(/\s+as\s+/i)[0];
                    } else {
                        selectAs.push(viewFlds[iVFlds]);
                    }

                    viewFlds[iVFlds] = utils.encodeNames(viewFlds[iVFlds]);     //=> SPENCODE
                    property = selectAs[iVFlds].replace(new RegExp(allowedTabCol.source + "\\_SEP\\_"), "");
                    //ERROR HANDLING - WARNING: 
                    if (that.debugMode && !warningShown) {
                        if (properties.indexOf(property) !== -1) {
                            warnAlias = property;
                        } else {
                            properties.push(property);
                        }
                        if (warnAlias.length !== 0) {
                            warningShown = true;
                            console.log("[crudeSP]: ReadItems - WARNING: Fieldname '" + warnAlias + "' occurs in more than one list, please specify an alias to distinguish the fields (n. W02)");
                        }
                    }
                    // /ERROR HANDLING - WARNING
                    if (viewFlds[iVFlds].match(new RegExp(allowedTabCol.source + "\\."))) {              //=> internal columname with alias  + _SEP_ + columnname     RGX: /^(.+?)\./
                        viewFlds[iVFlds] = viewFlds[iVFlds].replace(new RegExp("^(" + allowedTabCol.source + "?)\\."), function (match) {

                            return match[0] + "_SEP_";
                        });
                    }

                    if (tableRow.get_item(viewFlds[iVFlds]) !== null) { //sap-eternityDummy 9999-12-31 is displayed as null

                        if (tableRow.get_item(viewFlds[iVFlds]).constructor.getName() === "SP.FieldLookupValue") {
                            containerObj[property] = tableRow.get_item(viewFlds[iVFlds]).get_lookupValue();
                        } else {
                            containerObj[property] = tableRow.get_item(viewFlds[iVFlds]);
                        }
                    } else {
                        containerObj[property] = "";
                    }
                }

                if (!$.isEmptyObject(inlineProps)) {
                    for (var prop in inlineProps) {
                        if (inlineProps.hasOwnProperty(prop)) {
                            containerObj[prop] = inlineProps[prop];
                        }
                    }
                }
                outputArray.push(containerObj);
            }

            //trigger callback function
            callBack(outputArray);

        }

        //on error
        function listFail(sender, args) {
            if (typeof callBackFail !== "undefined") {
                callBackFail(args.get_message());
            }
        }
    };

    ///////////////////// II UpdateItems /////////////////////
    this.updateItems = function (callBack, callBackFail) {
        var listNEntities;
        var list;
        var listItems;
        var listItemsToPermit = [];
        var errorFound = false;

        utils.getContextWeb().then(function (data) {
            listColumns = data;
            listNEntities = new commons.GetListAndItems();
            list = listNEntities.list;
            listItems = listNEntities.listItems;

            if (typeof that.roles !== "undefined") {
                commons.checkIfUsersGroupsExist(2);
                doTheUpdate();
            } else {
                doTheUpdate();
            }

            function doTheUpdate() {
                context.load(listItems);
                context.executeQueryAsync(updatingList, updateFailed);
            }
        });

        //on success
        function updatingList() {
            commons.iteratorFn(listItems, processEntityUpdate);
            context.executeQueryAsync(updateSuccess, updateFailed);

            function processEntityUpdate(tableRow) {
                if (affectedRows > 0) {
                    listItemsToPermit.push(tableRow);
                    //ERROR HANDLING
                    if (that.debugMode) {
                        if (typeof that.set === "undefined" && typeof that.roles === "undefined") {
                            throw new spqException("UpdateItems: missing parameters", "no set or roles operation defined in update (n. 201)");
                        }
                    }
                    // /ERROR HANDLING

                    if (typeof that.set !== "undefined") {
                        //!!!DEPRECIATED!!!
                        if (typeof that.set === "string") {
                            that.set = commons.replaceCommas(that.set);
                            var toSet = that.set.split(","); ///[^\\],/ doesnt work, negative lookbehinds are not supported if js regex engine
                            for (var iS = 0, iSSum = toSet.length; iS < iSSum; iS++) {
                                //ERROR HANDLING
                                if (that.debugMode) {
                                    commons.findFieldErrors(toSet[iS], 2);
                                }
                                // /ERROR HANDLING

                                var sliceLeft = toSet[iS].split("=")[0].trim();
                                var sliceRight = toSet[iS].split("=")[1].trim();

                                sliceLeft = utils.encodeNames(sliceLeft);   //=> SPENCODE

                                //replace _COMMA_ with ,
                                sliceRight = sliceRight.replace(/\_COMMA\_/, ",");

                                //convert to ISO-date if necessary
                                if (commons.stringIsDateFLAWED(sliceRight)) {
                                    sliceRight = new Date(sliceRight).toISOString();
                                }
                                tableRow.set_item(utils.encodeNames(sliceLeft), sliceRight);       //=> SPENCODE
                            }
                            tableRow.update();
                            // /!!!DEPRECIATED!!!
                        } else if (typeof that.set === "object") {
                            //ERROR HANDLING
                            if (that.debugMode) {
                                if (typeof that.set.length !== "undefined") {
                                    throw new spqException("UpdateItems: invalid token", "set cannot be an array (n. 215)");
                                }
                            }
                            // /ERROR HANDLING
                            for (var key in that.set) {
                                if (that.set.hasOwnProperty(key)) {
                                    var sliceLeftObj = utils.encodeNames(key); //=> SPENCODE
                                    var sliceRightObj = that.set[key];
                                    //FEHLERBEHANLDUNG
                                    if (that.debugMode) {
                                        commons.findFieldErrors(key, 2, false);
                                    }
                                    // /ERROR HANDLING

                                    tableRow.set_item(utils.encodeNames(sliceLeftObj), sliceRightObj); //=> SPENCODE
                                }
                            }
                            tableRow.update();
                        } else {
                            if (that.debugMode) {
                                throw new spqException("UpdateItems: invalid token", "values must be of type string or object (n. 216)");
                            }
                        }
                    }
                }
            }
        }

        function updateFailed(sender, args) {
            if (typeof callBackFail !== "undefined") {
                callBackFail(args.get_message());
            }
        }

        function updateSuccess() {
            if (affectedRows > 0) {

                if (typeof that.roles !== "undefined") {
                    function deleteRolesPerRowWithPromise(listItemsArray) {
                        var deferred = new $.Deferred();
                        var copiedListItems = [];
                        for (var iLIA = 0, iLIASum = listItemsArray.length; iLIA < iLIASum; iLIA++) {
                            copiedListItems.push(listItemsArray[iLIA]); //arrays and other objects are addressed by reference, iterate and push into new array to prevent overwriting of original variables
                        }                                               

                        (function doThat() {
                            if (copiedListItems.length !== 0) {
                                utils.deleteRoles(copiedListItems[0]).then(function () {
                                    copiedListItems.splice(0, 1);
                                    doThat();
                                });
                            } else {
                                deferred.resolve();
                            }
                        })();

                        return deferred.promise();
                    }

                    var roleObjectOrArray = that.roles[0] || that.roles;

                    if (roleObjectOrArray.deleteExisting) {
                        deleteRolesPerRowWithPromise(listItemsToPermit).then(function () {
                            if (!errorFound) {
                                utils.assignRoles(listItemsToPermit, executeCallback);
                            }
                        });
                    } else {
                        utils.assignRoles(listItemsToPermit, executeCallback);
                    }
                }
                else {
                    executeCallback();
                }
            }
        }

        function executeCallback() {
            if (typeof callBack !== "undefined") {
                callBack(); //callback-Fn is optional
            }
        }

    };


    ///////////////////// DeleteItem /////////////////////
    this.deleteItems = function (callBack, callBackFail) {
        var listNEntities;
        var list;
        var listItems;
        var listItemIdToDelete = [];

        utils.getContextWeb().then(function (data) {
            listColumns = data;
            listNEntities = new commons.GetListAndItems();
            list = listNEntities.list;
            listItems = listNEntities.listItems;
            listItemIdToDelete = [];
            context.load(listItems);
            context.executeQueryAsync(deleting, deleteFailed);
        });

        //on success
        function deleting() {

            commons.iteratorFn(listItems, processEntityDelete);

            function processEntityDelete(tableRow) {
                if (affectedRows > 0) {
                    listItemIdToDelete.push(tableRow);
                }
            }

            if (affectedRows > 0) {
                for (var iTd = 0, toDeleteSum = listItemIdToDelete.length; iTd < toDeleteSum; iTd++) {
                    listItemIdToDelete[iTd].deleteObject();
                }
                context.executeQueryAsync(deleteSuccess, deleteFailed);
            } else {
                deleteSuccess();    //no rows affected no fatal error
            }
        }

        function deleteFailed(sender, args) {
            if (typeof callBackFail !== "undefined") {
                callBackFail(args.get_message());
            }
        }

        function deleteSuccess() {
            if (typeof callBack !== "undefined") {
                callBack();       //callback-Fn is optional
            }
        }
    };

    ///////////////////// Create /////////////////////
    this.createItems = function (callBack, callBackFail) {
        var listItemsToPermit = [];
        var list;
     
        utils.getContextWeb().then(function (data) {
            listColumns = data;
            list = web.get_lists().getByTitle(primaryList);
            if (typeof that.roles !== "undefined") {
                commons.checkIfUsersGroupsExist(4);
                doCreate();
            } else {
                doCreate();
            }
        });

        function doCreate() {
            //ERROR HANDLING
            if (that.debugMode) {
                if (typeof that.values === "undefined") {
                    throw new spqException("CreateItems: missing parameters", "no values entered (n. 401)");
                }
            }
            // /ERROR HANDLING

            var listItem;

            //!!!DEPRECIATED!!!
            if (typeof that.values === "string") {
                var values = that.values.split(/\)\s+?\,\s+?\(/); //var values = that.values.split("),(");
                for (var iV = 0, iVSum = values.length; iV < iVSum; iV++) {
                    var item = new SP.ListItemCreationInformation();
                    listItem = list.addItem(item);
                    listItemsToPermit.push(listItem);
                    values[iV] = values[iV].replace(/[\(\)]/g, "");

                    //ERROR HANDLING
                    if (that.debugMode) {
                        if (values[iV].match(/[\u00c4\u00e4\u00d6\u00f6\u00dc\u00fc\u00df\w\+#!*\/%\?\-\_\s]+,[\u00c4\u00e4\u00d6\u00f6\u00dc\u00fc\u00df\w\+#!*\/%\?\-\_\s]+,/)) { //=> /[\w\s]+,[\w\s]+,/
                            throw new spqException("CreateItems: invalid token", "unexpected comma in '" + values[iV] + "' (n. 415)");
                        }
                    }
                    // /ERROR HANDLING

                    values[iV] = commons.replaceCommas(values[iV]);
                    var columnValues = values[iV].split(","); 

                    for (var iCV = 0, iCVSum = columnValues.length; iCV < iCVSum; iCV++) {
                        columnValues[iCV] = columnValues[iCV].replace(/\\,/, ",");
                        //ERROR HANDLING
                        if (that.debugMode) {
                            commons.findFieldErrors(columnValues[iCV], 4);
                        }
                        // /ERROR HANDLING
                        var sliceLeft = columnValues[iCV].split("=")[0].trim();
                        var sliceRight = columnValues[iCV].split("=")[1].trim().replace(new RegExp(/\_COMMA\_/g), ",");
                        listItem.set_item(utils.encodeNames(sliceLeft), sliceRight);       //=>SPENCODE
                    }

                    listItem.update();

                    context.load(listItem);
                }
                // /!!!DEPRECIATED!!!
            } else if (typeof that.values === "object") {
                //do that with objects  

                var container = [];
                if (typeof that.values.length === "undefined") {
                    container.push(that.values);
                } else {
                    container = that.values;
                }

                for (var obj in container) {
                    var itemOb = new SP.ListItemCreationInformation();
                    listItem = list.addItem(itemOb);
                    listItemsToPermit.push(listItem);

                    for (var key in container[obj]) {
                        if (container[obj].hasOwnProperty(key)) {
                            //ERROR HANDLING
                            if (that.debugMode) {
                                commons.findFieldErrors(key, 4, false);
                            }
                            // /ERROR HANDLING
                            var sliceLeftObj = utils.encodeNames(key); //=>SPENCODE
                            var sliceRightObj = container[obj][key];

                            if (commons.stringIsDateFLAWED(sliceRightObj)) {
                                sliceRightObj = new Date(sliceRightObj).toISOString();
                            }

                            listItem.set_item(utils.encodeNames(sliceLeftObj), sliceRightObj); //=>SPENCODE
                        }
                    }
                    listItem.update();

                    context.load(listItem);
                }
            } else {
                if (that.debugMode) {
                    throw new spqException("CreateItems: invalid token", "values must be of type string or array (n. 416)");
                }
            }
            context.executeQueryAsync(createSuccess, createFailed);
        }

        function createFailed(sender, args) {
            if (typeof callBackFail !== "undefined") {
                callBackFail(args.get_message());
            }
        }

        function createSuccess() {
            if (typeof callBack !== "undefined") {
                callBack(); //callback-Fn is optional!
            }
            if (typeof that.roles !== "undefined") {
                utils.assignRoles(listItemsToPermit, callBack);
            }
        }
    };
    ///////////////////////////////// ##### /CRUD Methods ##### ///////////////////////////////////////

    ///////////////////////////////// ##### Upload To Library ##### /////////////////////////////////
    this.uploadToLibrary = function (callBack, callBackFail) {
        if (that.debugMode) {
            if (typeof that.file === "undefined") {
                throw new spqException("Upload to library: missing parameters", "no file provided (n. 501)");
            }
        }

        if (typeof that.site === "undefined") {
            that.site = window.location.href.substring(0, window.location.href.indexOf("SitePages/Home.aspx"));
        }
        getFileBuffer(that.file).then(function () {
            if (typeof that.filename === "undefined") {
                that.filename = that.file.name;
            }
            var endpoint = that.site + "/_api/web/GetFolderByServerRelativeUrl('" + that.list + "')/Files" +
                "/Add(url='" + that.filename + "', overwrite=true)?$expand=ListItemAllFields";
            var request = new XMLHttpRequest();
            request.open("POST", endpoint, true);
            request.setRequestHeader("Accept", "application/json; odata=verbose");
            request.setRequestHeader("X-RequestDigest", $("#__REQUESTDIGEST").val());
            //request.setRequestHeader("content-length", fileBytes.byteLength);     => refused to set unsafe header
            request.processData = false;
            var done = false;
            request.onreadystatechange = function (data) {
                if (request.readyState === 4 && !done) {
                    done = true;
                    if (typeof that.set !== "undefined") {
                        var listItemId = JSON.parse(data.currentTarget.response).d.ListItemAllFields.Id;
                        var setFieldsWhereFileUploaded = new csp.Operation({
                            set: that.set,
                            list: that.list,
                            where: listItemId
                        });
                        if (typeof that.roles !== "undefined") {
                            setFieldsWhereFileUploaded.roles = that.roles;
                        }
                        if (typeof that.list !== "undefined") {
                            setFieldsWhereFileUploaded.site = that.site;
                        }
                        setFieldsWhereFileUploaded.updateItems(callBack, callBackFail);
                    } else {
                        callBack(data);
                    }
                }
            };
            request.onerror = function (message) {
                callBackFail(message);
            };
            request.send(null);
        });

        function getFileBuffer(file) {
            //foreign code - Scott Hillier: http://www.shillier.com/archive/2013/03/26/uploading-files-in-sharepoint-2013-using-csom-and-rest.aspx 
            var deferredInner = $.Deferred();
            var reader = new FileReader();
            reader.onload = function (e) {
                deferredInner.resolve(e.target.result);
            }

            reader.onerror = function (e) {
                deferredInner.reject(e.target.error);
            }

            reader.readAsArrayBuffer(file);
            return deferredInner.promise();
        };
    }

    ///////////////////////////////// ##### /Upload To Library ##### /////////////////////////////////

    /////////////////////////////// /O P E R A T I O N //////////////////////////////
}
// /[crudeSP]


