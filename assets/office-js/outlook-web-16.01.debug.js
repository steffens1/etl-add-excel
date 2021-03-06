/* Outlook web application specific API library */
/* Version: 16.0.11527.30000 */
/*
/*!
Copyright (c) Microsoft Corporation.  All rights reserved.
*/
/*!
Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
This file also contains the following Promise implementation (with a few small modifications):
* @overview es6-promise - a tiny implementation of Promises/A+.
* @copyright Copyright (c) 2014 Yehuda Katz, Tom Dale, Stefan Penner and contributors (Conversion to ES6 API by Jake Archibald)
* @license   Licensed under MIT license
*            See https://raw.githubusercontent.com/jakearchibald/es6-promise/master/LICENSE
* @version   2.3.0
*/
var __extends = this && this.__extends || function(d, b)
    {
        for(var p in b)
            if(b.hasOwnProperty(p))
                d[p] = b[p];
        function __()
        {
            this.constructor = d
        }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype,new __)
    };
var OfficeExt;
(function(OfficeExt)
{
    var MicrosoftAjaxFactory = function()
        {
            function MicrosoftAjaxFactory(){}
            MicrosoftAjaxFactory.prototype.isMsAjaxLoaded = function()
            {
                if(typeof Sys !== "undefined" && typeof Type !== "undefined" && Sys.StringBuilder && typeof Sys.StringBuilder === "function" && Type.registerNamespace && typeof Type.registerNamespace === "function" && Type.registerClass && typeof Type.registerClass === "function" && typeof Function._validateParams === "function" && Sys.Serialization && Sys.Serialization.JavaScriptSerializer && typeof Sys.Serialization.JavaScriptSerializer.serialize === "function")
                    return true;
                else
                    return false
            };
            MicrosoftAjaxFactory.prototype.loadMsAjaxFull = function(callback)
            {
                var msAjaxCDNPath = (window.location.protocol.toLowerCase() === "https:" ? "https:" : "http:") + "//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js";
                OSF.OUtil.loadScript(msAjaxCDNPath,callback)
            };
            Object.defineProperty(MicrosoftAjaxFactory.prototype,"msAjaxError",{
                get: function()
                {
                    if(this._msAjaxError == null && this.isMsAjaxLoaded())
                        this._msAjaxError = Error;
                    return this._msAjaxError
                },
                set: function(errorClass)
                {
                    this._msAjaxError = errorClass
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(MicrosoftAjaxFactory.prototype,"msAjaxString",{
                get: function()
                {
                    if(this._msAjaxString == null && this.isMsAjaxLoaded())
                        this._msAjaxString = String;
                    return this._msAjaxString
                },
                set: function(stringClass)
                {
                    this._msAjaxString = stringClass
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(MicrosoftAjaxFactory.prototype,"msAjaxDebug",{
                get: function()
                {
                    if(this._msAjaxDebug == null && this.isMsAjaxLoaded())
                        this._msAjaxDebug = Sys.Debug;
                    return this._msAjaxDebug
                },
                set: function(debugClass)
                {
                    this._msAjaxDebug = debugClass
                },
                enumerable: true,
                configurable: true
            });
            return MicrosoftAjaxFactory
        }();
    OfficeExt.MicrosoftAjaxFactory = MicrosoftAjaxFactory
})(OfficeExt || (OfficeExt = {}));
var OsfMsAjaxFactory = new OfficeExt.MicrosoftAjaxFactory;
var OSF = OSF || {};
var OfficeExt;
(function(OfficeExt)
{
    var SafeStorage = function()
        {
            function SafeStorage(_internalStorage)
            {
                this._internalStorage = _internalStorage
            }
            SafeStorage.prototype.getItem = function(key)
            {
                try
                {
                    return this._internalStorage && this._internalStorage.getItem(key)
                }
                catch(e)
                {
                    return null
                }
            };
            SafeStorage.prototype.setItem = function(key, data)
            {
                try
                {
                    this._internalStorage && this._internalStorage.setItem(key,data)
                }
                catch(e){}
            };
            SafeStorage.prototype.clear = function()
            {
                try
                {
                    this._internalStorage && this._internalStorage.clear()
                }
                catch(e){}
            };
            SafeStorage.prototype.removeItem = function(key)
            {
                try
                {
                    this._internalStorage && this._internalStorage.removeItem(key)
                }
                catch(e){}
            };
            SafeStorage.prototype.getKeysWithPrefix = function(keyPrefix)
            {
                var keyList = [];
                try
                {
                    var len = this._internalStorage && this._internalStorage.length || 0;
                    for(var i = 0; i < len; i++)
                    {
                        var key = this._internalStorage.key(i);
                        if(key.indexOf(keyPrefix) === 0)
                            keyList.push(key)
                    }
                }
                catch(e){}
                return keyList
            };
            return SafeStorage
        }();
    OfficeExt.SafeStorage = SafeStorage
})(OfficeExt || (OfficeExt = {}));
OSF.XdmFieldName = {
    ConversationUrl: "ConversationUrl",
    AppId: "AppId"
};
OSF.WindowNameItemKeys = {
    BaseFrameName: "baseFrameName",
    HostInfo: "hostInfo",
    XdmInfo: "xdmInfo",
    SerializerVersion: "serializerVersion",
    AppContext: "appContext"
};
OSF.OUtil = function()
{
    var _uniqueId = -1;
    var _xdmInfoKey = "&_xdm_Info=";
    var _serializerVersionKey = "&_serializer_version=";
    var _xdmSessionKeyPrefix = "_xdm_";
    var _serializerVersionKeyPrefix = "_serializer_version=";
    var _fragmentSeparator = "#";
    var _fragmentInfoDelimiter = "&";
    var _classN = "class";
    var _loadedScripts = {};
    var _defaultScriptLoadingTimeout = 3e4;
    var _safeSessionStorage = null;
    var _safeLocalStorage = null;
    var _rndentropy = (new Date).getTime();
    function _random()
    {
        var nextrand = 2147483647 * Math.random();
        nextrand ^= _rndentropy ^ (new Date).getMilliseconds() << Math.floor(Math.random() * (31 - 10));
        return nextrand.toString(16)
    }
    function _getSessionStorage()
    {
        if(!_safeSessionStorage)
        {
            try
            {
                var sessionStorage = window.sessionStorage
            }
            catch(ex)
            {
                sessionStorage = null
            }
            _safeSessionStorage = new OfficeExt.SafeStorage(sessionStorage)
        }
        return _safeSessionStorage
    }
    function _reOrderTabbableElements(elements)
    {
        var bucket0 = [];
        var bucketPositive = [];
        var i;
        var len = elements.length;
        var ele;
        for(i = 0; i < len; i++)
        {
            ele = elements[i];
            if(ele.tabIndex)
            {
                if(ele.tabIndex > 0)
                    bucketPositive.push(ele);
                else if(ele.tabIndex === 0)
                    bucket0.push(ele)
            }
            else
                bucket0.push(ele)
        }
        bucketPositive = bucketPositive.sort(function(left, right)
        {
            var diff = left.tabIndex - right.tabIndex;
            if(diff === 0)
                diff = bucketPositive.indexOf(left) - bucketPositive.indexOf(right);
            return diff
        });
        return[].concat(bucketPositive,bucket0)
    }
    return{
            set_entropy: function OSF_OUtil$set_entropy(entropy)
            {
                if(typeof entropy == "string")
                    for(var i = 0; i < entropy.length; i += 4)
                    {
                        var temp = 0;
                        for(var j = 0; j < 4 && i + j < entropy.length; j++)
                            temp = (temp << 8) + entropy.charCodeAt(i + j);
                        _rndentropy ^= temp
                    }
                else if(typeof entropy == "number")
                    _rndentropy ^= entropy;
                else
                    _rndentropy ^= 2147483647 * Math.random();
                _rndentropy &= 2147483647
            },
            extend: function OSF_OUtil$extend(child, parent)
            {
                var F = function(){};
                F.prototype = parent.prototype;
                child.prototype = new F;
                child.prototype.constructor = child;
                child.uber = parent.prototype;
                if(parent.prototype.constructor === Object.prototype.constructor)
                    parent.prototype.constructor = parent
            },
            setNamespace: function OSF_OUtil$setNamespace(name, parent)
            {
                if(parent && name && !parent[name])
                    parent[name] = {}
            },
            unsetNamespace: function OSF_OUtil$unsetNamespace(name, parent)
            {
                if(parent && name && parent[name])
                    delete parent[name]
            },
            serializeSettings: function OSF_OUtil$serializeSettings(settingsCollection)
            {
                var ret = {};
                for(var key in settingsCollection)
                {
                    var value = settingsCollection[key];
                    try
                    {
                        if(JSON)
                            value = JSON.stringify(value,function dateReplacer(k, v)
                            {
                                return OSF.OUtil.isDate(this[k]) ? OSF.DDA.SettingsManager.DateJSONPrefix + this[k].getTime() + OSF.DDA.SettingsManager.DataJSONSuffix : v
                            });
                        else
                            value = Sys.Serialization.JavaScriptSerializer.serialize(value);
                        ret[key] = value
                    }
                    catch(ex){}
                }
                return ret
            },
            deserializeSettings: function OSF_OUtil$deserializeSettings(serializedSettings)
            {
                var ret = {};
                serializedSettings = serializedSettings || {};
                for(var key in serializedSettings)
                {
                    var value = serializedSettings[key];
                    try
                    {
                        if(JSON)
                            value = JSON.parse(value,function dateReviver(k, v)
                            {
                                var d;
                                if(typeof v === "string" && v && v.length > 6 && v.slice(0,5) === OSF.DDA.SettingsManager.DateJSONPrefix && v.slice(-1) === OSF.DDA.SettingsManager.DataJSONSuffix)
                                {
                                    d = new Date(parseInt(v.slice(5,-1)));
                                    if(d)
                                        return d
                                }
                                return v
                            });
                        else
                            value = Sys.Serialization.JavaScriptSerializer.deserialize(value,true);
                        ret[key] = value
                    }
                    catch(ex){}
                }
                return ret
            },
            loadScript: function OSF_OUtil$loadScript(url, callback, timeoutInMs)
            {
                if(url && callback)
                {
                    var doc = window.document;
                    var _loadedScriptEntry = _loadedScripts[url];
                    if(!_loadedScriptEntry)
                    {
                        var script = doc.createElement("script");
                        script.type = "text/javascript";
                        _loadedScriptEntry = {
                            loaded: false,
                            pendingCallbacks: [callback],
                            timer: null
                        };
                        _loadedScripts[url] = _loadedScriptEntry;
                        var onLoadCallback = function OSF_OUtil_loadScript$onLoadCallback()
                            {
                                if(_loadedScriptEntry.timer != null)
                                {
                                    clearTimeout(_loadedScriptEntry.timer);
                                    delete _loadedScriptEntry.timer
                                }
                                _loadedScriptEntry.loaded = true;
                                var pendingCallbackCount = _loadedScriptEntry.pendingCallbacks.length;
                                for(var i = 0; i < pendingCallbackCount; i++)
                                {
                                    var currentCallback = _loadedScriptEntry.pendingCallbacks.shift();
                                    currentCallback()
                                }
                            };
                        var onLoadError = function OSF_OUtil_loadScript$onLoadError()
                            {
                                delete _loadedScripts[url];
                                if(_loadedScriptEntry.timer != null)
                                {
                                    clearTimeout(_loadedScriptEntry.timer);
                                    delete _loadedScriptEntry.timer
                                }
                                var pendingCallbackCount = _loadedScriptEntry.pendingCallbacks.length;
                                for(var i = 0; i < pendingCallbackCount; i++)
                                {
                                    var currentCallback = _loadedScriptEntry.pendingCallbacks.shift();
                                    currentCallback()
                                }
                            };
                        if(script.readyState)
                            script.onreadystatechange = function()
                            {
                                if(script.readyState == "loaded" || script.readyState == "complete")
                                {
                                    script.onreadystatechange = null;
                                    onLoadCallback()
                                }
                            };
                        else
                            script.onload = onLoadCallback;
                        script.onerror = onLoadError;
                        timeoutInMs = timeoutInMs || _defaultScriptLoadingTimeout;
                        _loadedScriptEntry.timer = setTimeout(onLoadError,timeoutInMs);
                        script.setAttribute("crossOrigin","anonymous");
                        script.src = url;
                        doc.getElementsByTagName("head")[0].appendChild(script)
                    }
                    else if(_loadedScriptEntry.loaded)
                        callback();
                    else
                        _loadedScriptEntry.pendingCallbacks.push(callback)
                }
            },
            loadCSS: function OSF_OUtil$loadCSS(url)
            {
                if(url)
                {
                    var doc = window.document;
                    var link = doc.createElement("link");
                    link.type = "text/css";
                    link.rel = "stylesheet";
                    link.href = url;
                    doc.getElementsByTagName("head")[0].appendChild(link)
                }
            },
            parseEnum: function OSF_OUtil$parseEnum(str, enumObject)
            {
                var parsed = enumObject[str.trim()];
                if(typeof parsed == "undefined")
                {
                    OsfMsAjaxFactory.msAjaxDebug.trace("invalid enumeration string:" + str);
                    throw OsfMsAjaxFactory.msAjaxError.argument("str");
                }
                return parsed
            },
            delayExecutionAndCache: function OSF_OUtil$delayExecutionAndCache()
            {
                var obj = {calc: arguments[0]};
                return function()
                    {
                        if(obj.calc)
                        {
                            obj.val = obj.calc.apply(this,arguments);
                            delete obj.calc
                        }
                        return obj.val
                    }
            },
            getUniqueId: function OSF_OUtil$getUniqueId()
            {
                _uniqueId = _uniqueId + 1;
                return _uniqueId.toString()
            },
            formatString: function OSF_OUtil$formatString()
            {
                var args = arguments;
                var source = args[0];
                return source.replace(/{(\d+)}/gm,function(match, number)
                    {
                        var index = parseInt(number,10) + 1;
                        return args[index] === undefined ? "{" + number + "}" : args[index]
                    })
            },
            generateConversationId: function OSF_OUtil$generateConversationId()
            {
                return[_random(),_random(),(new Date).getTime().toString()].join("_")
            },
            getFrameName: function OSF_OUtil$getFrameName(cacheKey)
            {
                return _xdmSessionKeyPrefix + cacheKey + this.generateConversationId()
            },
            addXdmInfoAsHash: function OSF_OUtil$addXdmInfoAsHash(url, xdmInfoValue)
            {
                return OSF.OUtil.addInfoAsHash(url,_xdmInfoKey,xdmInfoValue,false)
            },
            addSerializerVersionAsHash: function OSF_OUtil$addSerializerVersionAsHash(url, serializerVersion)
            {
                return OSF.OUtil.addInfoAsHash(url,_serializerVersionKey,serializerVersion,true)
            },
            addInfoAsHash: function OSF_OUtil$addInfoAsHash(url, keyName, infoValue, encodeInfo)
            {
                url = url.trim() || "";
                var urlParts = url.split(_fragmentSeparator);
                var urlWithoutFragment = urlParts.shift();
                var fragment = urlParts.join(_fragmentSeparator);
                var newFragment;
                if(encodeInfo)
                    newFragment = [keyName,encodeURIComponent(infoValue),fragment].join("");
                else
                    newFragment = [fragment,keyName,infoValue].join("");
                return[urlWithoutFragment,_fragmentSeparator,newFragment].join("")
            },
            parseHostInfoFromWindowName: function OSF_OUtil$parseHostInfoFromWindowName(skipSessionStorage, windowName)
            {
                return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage,windowName,OSF.WindowNameItemKeys.HostInfo)
            },
            parseXdmInfo: function OSF_OUtil$parseXdmInfo(skipSessionStorage)
            {
                var xdmInfoValue = OSF.OUtil.parseXdmInfoWithGivenFragment(skipSessionStorage,window.location.hash);
                if(!xdmInfoValue)
                    xdmInfoValue = OSF.OUtil.parseXdmInfoFromWindowName(skipSessionStorage,window.name);
                return xdmInfoValue
            },
            parseXdmInfoFromWindowName: function OSF_OUtil$parseXdmInfoFromWindowName(skipSessionStorage, windowName)
            {
                return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage,windowName,OSF.WindowNameItemKeys.XdmInfo)
            },
            parseXdmInfoWithGivenFragment: function OSF_OUtil$parseXdmInfoWithGivenFragment(skipSessionStorage, fragment)
            {
                return OSF.OUtil.parseInfoWithGivenFragment(_xdmInfoKey,_xdmSessionKeyPrefix,false,skipSessionStorage,fragment)
            },
            parseSerializerVersion: function OSF_OUtil$parseSerializerVersion(skipSessionStorage)
            {
                var serializerVersion = OSF.OUtil.parseSerializerVersionWithGivenFragment(skipSessionStorage,window.location.hash);
                if(isNaN(serializerVersion))
                    serializerVersion = OSF.OUtil.parseSerializerVersionFromWindowName(skipSessionStorage,window.name);
                return serializerVersion
            },
            parseSerializerVersionFromWindowName: function OSF_OUtil$parseSerializerVersionFromWindowName(skipSessionStorage, windowName)
            {
                return parseInt(OSF.OUtil.parseInfoFromWindowName(skipSessionStorage,windowName,OSF.WindowNameItemKeys.SerializerVersion))
            },
            parseSerializerVersionWithGivenFragment: function OSF_OUtil$parseSerializerVersionWithGivenFragment(skipSessionStorage, fragment)
            {
                return parseInt(OSF.OUtil.parseInfoWithGivenFragment(_serializerVersionKey,_serializerVersionKeyPrefix,true,skipSessionStorage,fragment))
            },
            parseInfoFromWindowName: function OSF_OUtil$parseInfoFromWindowName(skipSessionStorage, windowName, infoKey)
            {
                try
                {
                    var windowNameObj = JSON.parse(windowName);
                    var infoValue = windowNameObj != null ? windowNameObj[infoKey] : null;
                    var osfSessionStorage = _getSessionStorage();
                    if(!skipSessionStorage && osfSessionStorage && windowNameObj != null)
                    {
                        var sessionKey = windowNameObj[OSF.WindowNameItemKeys.BaseFrameName] + infoKey;
                        if(infoValue)
                            osfSessionStorage.setItem(sessionKey,infoValue);
                        else
                            infoValue = osfSessionStorage.getItem(sessionKey)
                    }
                    return infoValue
                }
                catch(Exception)
                {
                    return null
                }
            },
            parseInfoWithGivenFragment: function OSF_OUtil$parseInfoWithGivenFragment(infoKey, infoKeyPrefix, decodeInfo, skipSessionStorage, fragment)
            {
                var fragmentParts = fragment.split(infoKey);
                var infoValue = fragmentParts.length > 1 ? fragmentParts[fragmentParts.length - 1] : null;
                if(decodeInfo && infoValue != null)
                {
                    if(infoValue.indexOf(_fragmentInfoDelimiter) >= 0)
                        infoValue = infoValue.split(_fragmentInfoDelimiter)[0];
                    infoValue = decodeURIComponent(infoValue)
                }
                var osfSessionStorage = _getSessionStorage();
                if(!skipSessionStorage && osfSessionStorage)
                {
                    var sessionKeyStart = window.name.indexOf(infoKeyPrefix);
                    if(sessionKeyStart > -1)
                    {
                        var sessionKeyEnd = window.name.indexOf(";",sessionKeyStart);
                        if(sessionKeyEnd == -1)
                            sessionKeyEnd = window.name.length;
                        var sessionKey = window.name.substring(sessionKeyStart,sessionKeyEnd);
                        if(infoValue)
                            osfSessionStorage.setItem(sessionKey,infoValue);
                        else
                            infoValue = osfSessionStorage.getItem(sessionKey)
                    }
                }
                return infoValue
            },
            getConversationId: function OSF_OUtil$getConversationId()
            {
                var searchString = window.location.search;
                var conversationId = null;
                if(searchString)
                {
                    var index = searchString.indexOf("&");
                    conversationId = index > 0 ? searchString.substring(1,index) : searchString.substr(1);
                    if(conversationId && conversationId.charAt(conversationId.length - 1) === "=")
                    {
                        conversationId = conversationId.substring(0,conversationId.length - 1);
                        if(conversationId)
                            conversationId = decodeURIComponent(conversationId)
                    }
                }
                return conversationId
            },
            getInfoItems: function OSF_OUtil$getInfoItems(strInfo)
            {
                var items = strInfo.split("$");
                if(typeof items[1] == "undefined")
                    items = strInfo.split("|");
                if(typeof items[1] == "undefined")
                    items = strInfo.split("%7C");
                return items
            },
            getXdmFieldValue: function OSF_OUtil$getXdmFieldValue(xdmFieldName, skipSessionStorage)
            {
                var fieldValue = "";
                var xdmInfoValue = OSF.OUtil.parseXdmInfo(skipSessionStorage);
                if(xdmInfoValue)
                {
                    var items = OSF.OUtil.getInfoItems(xdmInfoValue);
                    if(items != undefined && items.length >= 3)
                        switch(xdmFieldName)
                        {
                            case OSF.XdmFieldName.ConversationUrl:
                                fieldValue = items[2];
                                break;
                            case OSF.XdmFieldName.AppId:
                                fieldValue = items[1];
                                break
                        }
                }
                return fieldValue
            },
            validateParamObject: function OSF_OUtil$validateParamObject(params, expectedProperties, callback)
            {
                var e = Function._validateParams(arguments,[{
                            name: "params",
                            type: Object,
                            mayBeNull: false
                        },{
                            name: "expectedProperties",
                            type: Object,
                            mayBeNull: false
                        },{
                            name: "callback",
                            type: Function,
                            mayBeNull: true
                        }]);
                if(e)
                    throw e;
                for(var p in expectedProperties)
                {
                    e = Function._validateParameter(params[p],expectedProperties[p],p);
                    if(e)
                        throw e;
                }
            },
            writeProfilerMark: function OSF_OUtil$writeProfilerMark(text)
            {
                if(window.msWriteProfilerMark)
                {
                    window.msWriteProfilerMark(text);
                    OsfMsAjaxFactory.msAjaxDebug.trace(text)
                }
            },
            outputDebug: function OSF_OUtil$outputDebug(text)
            {
                if(typeof OsfMsAjaxFactory !== "undefined" && OsfMsAjaxFactory.msAjaxDebug && OsfMsAjaxFactory.msAjaxDebug.trace)
                    OsfMsAjaxFactory.msAjaxDebug.trace(text)
            },
            defineNondefaultProperty: function OSF_OUtil$defineNondefaultProperty(obj, prop, descriptor, attributes)
            {
                descriptor = descriptor || {};
                for(var nd in attributes)
                {
                    var attribute = attributes[nd];
                    if(descriptor[attribute] == undefined)
                        descriptor[attribute] = true
                }
                Object.defineProperty(obj,prop,descriptor);
                return obj
            },
            defineNondefaultProperties: function OSF_OUtil$defineNondefaultProperties(obj, descriptors, attributes)
            {
                descriptors = descriptors || {};
                for(var prop in descriptors)
                    OSF.OUtil.defineNondefaultProperty(obj,prop,descriptors[prop],attributes);
                return obj
            },
            defineEnumerableProperty: function OSF_OUtil$defineEnumerableProperty(obj, prop, descriptor)
            {
                return OSF.OUtil.defineNondefaultProperty(obj,prop,descriptor,["enumerable"])
            },
            defineEnumerableProperties: function OSF_OUtil$defineEnumerableProperties(obj, descriptors)
            {
                return OSF.OUtil.defineNondefaultProperties(obj,descriptors,["enumerable"])
            },
            defineMutableProperty: function OSF_OUtil$defineMutableProperty(obj, prop, descriptor)
            {
                return OSF.OUtil.defineNondefaultProperty(obj,prop,descriptor,["writable","enumerable","configurable"])
            },
            defineMutableProperties: function OSF_OUtil$defineMutableProperties(obj, descriptors)
            {
                return OSF.OUtil.defineNondefaultProperties(obj,descriptors,["writable","enumerable","configurable"])
            },
            finalizeProperties: function OSF_OUtil$finalizeProperties(obj, descriptor)
            {
                descriptor = descriptor || {};
                var props = Object.getOwnPropertyNames(obj);
                var propsLength = props.length;
                for(var i = 0; i < propsLength; i++)
                {
                    var prop = props[i];
                    var desc = Object.getOwnPropertyDescriptor(obj,prop);
                    if(!desc.get && !desc.set)
                        desc.writable = descriptor.writable || false;
                    desc.configurable = descriptor.configurable || false;
                    desc.enumerable = descriptor.enumerable || true;
                    Object.defineProperty(obj,prop,desc)
                }
                return obj
            },
            mapList: function OSF_OUtil$MapList(list, mapFunction)
            {
                var ret = [];
                if(list)
                    for(var item in list)
                        ret.push(mapFunction(list[item]));
                return ret
            },
            listContainsKey: function OSF_OUtil$listContainsKey(list, key)
            {
                for(var item in list)
                    if(key == item)
                        return true;
                return false
            },
            listContainsValue: function OSF_OUtil$listContainsElement(list, value)
            {
                for(var item in list)
                    if(value == list[item])
                        return true;
                return false
            },
            augmentList: function OSF_OUtil$augmentList(list, addenda)
            {
                var add = list.push ? function(key, value)
                    {
                        list.push(value)
                    } : function(key, value)
                    {
                        list[key] = value
                    };
                for(var key in addenda)
                    add(key,addenda[key])
            },
            redefineList: function OSF_Outil$redefineList(oldList, newList)
            {
                for(var key1 in oldList)
                    delete oldList[key1];
                for(var key2 in newList)
                    oldList[key2] = newList[key2]
            },
            isArray: function OSF_OUtil$isArray(obj)
            {
                return Object.prototype.toString.apply(obj) === "[object Array]"
            },
            isFunction: function OSF_OUtil$isFunction(obj)
            {
                return Object.prototype.toString.apply(obj) === "[object Function]"
            },
            isDate: function OSF_OUtil$isDate(obj)
            {
                return Object.prototype.toString.apply(obj) === "[object Date]"
            },
            addEventListener: function OSF_OUtil$addEventListener(element, eventName, listener)
            {
                if(element.addEventListener)
                    element.addEventListener(eventName,listener,false);
                else if(Sys.Browser.agent === Sys.Browser.InternetExplorer && element.attachEvent)
                    element.attachEvent("on" + eventName,listener);
                else
                    element["on" + eventName] = listener
            },
            removeEventListener: function OSF_OUtil$removeEventListener(element, eventName, listener)
            {
                if(element.removeEventListener)
                    element.removeEventListener(eventName,listener,false);
                else if(Sys.Browser.agent === Sys.Browser.InternetExplorer && element.detachEvent)
                    element.detachEvent("on" + eventName,listener);
                else
                    element["on" + eventName] = null
            },
            getCookieValue: function OSF_OUtil$getCookieValue(cookieName)
            {
                var tmpCookieString = RegExp(cookieName + "[^;]+").exec(document.cookie);
                return tmpCookieString.toString().replace(/^[^=]+./,"")
            },
            xhrGet: function OSF_OUtil$xhrGet(url, onSuccess, onError)
            {
                var xmlhttp;
                try
                {
                    xmlhttp = new XMLHttpRequest;
                    xmlhttp.onreadystatechange = function()
                    {
                        if(xmlhttp.readyState == 4)
                            if(xmlhttp.status == 200)
                                onSuccess(xmlhttp.responseText);
                            else
                                onError(xmlhttp.status)
                    };
                    xmlhttp.open("GET",url,true);
                    xmlhttp.send()
                }
                catch(ex)
                {
                    onError(ex)
                }
            },
            xhrGetFull: function OSF_OUtil$xhrGetFull(url, oneDriveFileName, onSuccess, onError)
            {
                var xmlhttp;
                var requestedFileName = oneDriveFileName;
                try
                {
                    xmlhttp = new XMLHttpRequest;
                    xmlhttp.onreadystatechange = function()
                    {
                        if(xmlhttp.readyState == 4)
                            if(xmlhttp.status == 200)
                                onSuccess(xmlhttp,requestedFileName);
                            else
                                onError(xmlhttp.status)
                    };
                    xmlhttp.open("GET",url,true);
                    xmlhttp.send()
                }
                catch(ex)
                {
                    onError(ex)
                }
            },
            encodeBase64: function OSF_Outil$encodeBase64(input)
            {
                if(!input)
                    return input;
                var codex = "ABCDEFGHIJKLMNOP" + "QRSTUVWXYZabcdef" + "ghijklmnopqrstuv" + "wxyz0123456789+/=";
                var output = [];
                var temp = [];
                var index = 0;
                var c1,
                    c2,
                    c3,
                    a,
                    b,
                    c;
                var i;
                var length = input.length;
                do
                {
                    c1 = input.charCodeAt(index++);
                    c2 = input.charCodeAt(index++);
                    c3 = input.charCodeAt(index++);
                    i = 0;
                    a = c1 & 255;
                    b = c1 >> 8;
                    c = c2 & 255;
                    temp[i++] = a >> 2;
                    temp[i++] = (a & 3) << 4 | b >> 4;
                    temp[i++] = (b & 15) << 2 | c >> 6;
                    temp[i++] = c & 63;
                    if(!isNaN(c2))
                    {
                        a = c2 >> 8;
                        b = c3 & 255;
                        c = c3 >> 8;
                        temp[i++] = a >> 2;
                        temp[i++] = (a & 3) << 4 | b >> 4;
                        temp[i++] = (b & 15) << 2 | c >> 6;
                        temp[i++] = c & 63
                    }
                    if(isNaN(c2))
                        temp[i - 1] = 64;
                    else if(isNaN(c3))
                    {
                        temp[i - 2] = 64;
                        temp[i - 1] = 64
                    }
                    for(var t = 0; t < i; t++)
                        output.push(codex.charAt(temp[t]))
                } while(index < length);
                return output.join("")
            },
            getSessionStorage: function OSF_Outil$getSessionStorage()
            {
                return _getSessionStorage()
            },
            getLocalStorage: function OSF_Outil$getLocalStorage()
            {
                if(!_safeLocalStorage)
                {
                    try
                    {
                        var localStorage = window.localStorage
                    }
                    catch(ex)
                    {
                        localStorage = null
                    }
                    _safeLocalStorage = new OfficeExt.SafeStorage(localStorage)
                }
                return _safeLocalStorage
            },
            convertIntToCssHexColor: function OSF_Outil$convertIntToCssHexColor(val)
            {
                var hex = "#" + (Number(val) + 16777216).toString(16).slice(-6);
                return hex
            },
            attachClickHandler: function OSF_Outil$attachClickHandler(element, handler)
            {
                element.onclick = function(e)
                {
                    handler()
                };
                element.ontouchend = function(e)
                {
                    handler();
                    e.preventDefault()
                }
            },
            getQueryStringParamValue: function OSF_Outil$getQueryStringParamValue(queryString, paramName)
            {
                var e = Function._validateParams(arguments,[{
                            name: "queryString",
                            type: String,
                            mayBeNull: false
                        },{
                            name: "paramName",
                            type: String,
                            mayBeNull: false
                        }]);
                if(e)
                {
                    OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: Parameters cannot be null.");
                    return""
                }
                var queryExp = new RegExp("[\\?&]" + paramName + "=([^&#]*)","i");
                if(!queryExp.test(queryString))
                {
                    OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: The parameter is not found.");
                    return""
                }
                return queryExp.exec(queryString)[1]
            },
            isiOS: function OSF_Outil$isiOS()
            {
                return window.navigator.userAgent.match(/(iPad|iPhone|iPod)/g) ? true : false
            },
            isChrome: function OSF_Outil$isChrome()
            {
                return window.navigator.userAgent.indexOf("Chrome") > 0 && !OSF.OUtil.isEdge()
            },
            isEdge: function OSF_Outil$isEdge()
            {
                return window.navigator.userAgent.indexOf("Edge") > 0
            },
            isIE: function OSF_Outil$isIE()
            {
                return window.navigator.userAgent.indexOf("Trident") > 0
            },
            isFirefox: function OSF_Outil$isFirefox()
            {
                return window.navigator.userAgent.indexOf("Firefox") > 0
            },
            shallowCopy: function OSF_Outil$shallowCopy(sourceObj)
            {
                if(sourceObj == null)
                    return null;
                else if(!(sourceObj instanceof Object))
                    return sourceObj;
                else if(Array.isArray(sourceObj))
                {
                    var copyArr = [];
                    for(var i = 0; i < sourceObj.length; i++)
                        copyArr.push(sourceObj[i]);
                    return copyArr
                }
                else
                {
                    var copyObj = sourceObj.constructor();
                    for(var property in sourceObj)
                        if(sourceObj.hasOwnProperty(property))
                            copyObj[property] = sourceObj[property];
                    return copyObj
                }
            },
            createObject: function OSF_Outil$createObject(properties)
            {
                var obj = null;
                if(properties)
                {
                    obj = {};
                    var len = properties.length;
                    for(var i = 0; i < len; i++)
                        obj[properties[i].name] = properties[i].value
                }
                return obj
            },
            addClass: function OSF_OUtil$addClass(elmt, val)
            {
                if(!OSF.OUtil.hasClass(elmt,val))
                {
                    var className = elmt.getAttribute(_classN);
                    if(className)
                        elmt.setAttribute(_classN,className + " " + val);
                    else
                        elmt.setAttribute(_classN,val)
                }
            },
            removeClass: function OSF_OUtil$removeClass(elmt, val)
            {
                if(OSF.OUtil.hasClass(elmt,val))
                {
                    var className = elmt.getAttribute(_classN);
                    var reg = new RegExp("(\\s|^)" + val + "(\\s|$)");
                    className = className.replace(reg,"");
                    elmt.setAttribute(_classN,className)
                }
            },
            hasClass: function OSF_OUtil$hasClass(elmt, clsName)
            {
                var className = elmt.getAttribute(_classN);
                return className && className.match(new RegExp("(\\s|^)" + clsName + "(\\s|$)"))
            },
            focusToFirstTabbable: function OSF_OUtil$focusToFirstTabbable(all, backward)
            {
                var next;
                var focused = false;
                var candidate;
                var setFlag = function(e)
                    {
                        focused = true
                    };
                var findNextPos = function(allLen, currPos, backward)
                    {
                        if(currPos < 0 || currPos > allLen)
                            return-1;
                        else if(currPos === 0 && backward)
                            return-1;
                        else if(currPos === allLen - 1 && !backward)
                            return-1;
                        if(backward)
                            return currPos - 1;
                        else
                            return currPos + 1
                    };
                all = _reOrderTabbableElements(all);
                next = backward ? all.length - 1 : 0;
                if(all.length === 0)
                    return null;
                while(!focused && next >= 0 && next < all.length)
                {
                    candidate = all[next];
                    window.focus();
                    candidate.addEventListener("focus",setFlag);
                    candidate.focus();
                    candidate.removeEventListener("focus",setFlag);
                    next = findNextPos(all.length,next,backward);
                    if(!focused && candidate === document.activeElement)
                        focused = true
                }
                if(focused)
                    return candidate;
                else
                    return null
            },
            focusToNextTabbable: function OSF_OUtil$focusToNextTabbable(all, curr, shift)
            {
                var currPos;
                var next;
                var focused = false;
                var candidate;
                var setFlag = function(e)
                    {
                        focused = true
                    };
                var findCurrPos = function(all, curr)
                    {
                        var i = 0;
                        for(; i < all.length; i++)
                            if(all[i] === curr)
                                return i;
                        return-1
                    };
                var findNextPos = function(allLen, currPos, shift)
                    {
                        if(currPos < 0 || currPos > allLen)
                            return-1;
                        else if(currPos === 0 && shift)
                            return-1;
                        else if(currPos === allLen - 1 && !shift)
                            return-1;
                        if(shift)
                            return currPos - 1;
                        else
                            return currPos + 1
                    };
                all = _reOrderTabbableElements(all);
                currPos = findCurrPos(all,curr);
                next = findNextPos(all.length,currPos,shift);
                if(next < 0)
                    return null;
                while(!focused && next >= 0 && next < all.length)
                {
                    candidate = all[next];
                    candidate.addEventListener("focus",setFlag);
                    candidate.focus();
                    candidate.removeEventListener("focus",setFlag);
                    next = findNextPos(all.length,next,shift);
                    if(!focused && candidate === document.activeElement)
                        focused = true
                }
                if(focused)
                    return candidate;
                else
                    return null
            }
        }
}();
OSF.OUtil.Guid = function()
{
    var hexCode = ["0","1","2","3","4","5","6","7","8","9","a","b","c","d","e","f"];
    return{generateNewGuid: function OSF_Outil_Guid$generateNewGuid()
            {
                var result = "";
                var tick = (new Date).getTime();
                var index = 0;
                for(; index < 32 && tick > 0; index++)
                {
                    if(index == 8 || index == 12 || index == 16 || index == 20)
                        result += "-";
                    result += hexCode[tick % 16];
                    tick = Math.floor(tick / 16)
                }
                for(; index < 32; index++)
                {
                    if(index == 8 || index == 12 || index == 16 || index == 20)
                        result += "-";
                    result += hexCode[Math.floor(Math.random() * 16)]
                }
                return result
            }}
}();
window.OSF = OSF;
OSF.OUtil.setNamespace("OSF",window);
OSF.MessageIDs = {
    FetchBundleUrl: 0,
    LoadReactBundle: 1,
    LoadBundleSuccess: 2,
    LoadBundleError: 3
};
OSF.AppName = {
    Unsupported: 0,
    Excel: 1,
    Word: 2,
    PowerPoint: 4,
    Outlook: 8,
    ExcelWebApp: 16,
    WordWebApp: 32,
    OutlookWebApp: 64,
    Project: 128,
    AccessWebApp: 256,
    PowerpointWebApp: 512,
    ExcelIOS: 1024,
    Sway: 2048,
    WordIOS: 4096,
    PowerPointIOS: 8192,
    Access: 16384,
    Lync: 32768,
    OutlookIOS: 65536,
    OneNoteWebApp: 131072,
    OneNote: 262144,
    ExcelWinRT: 524288,
    WordWinRT: 1048576,
    PowerpointWinRT: 2097152,
    OutlookAndroid: 4194304,
    OneNoteWinRT: 8388608,
    ExcelAndroid: 8388609,
    VisioWebApp: 8388610,
    OneNoteIOS: 8388611,
    WordAndroid: 8388613,
    PowerpointAndroid: 8388614,
    Visio: 8388615,
    OneNoteAndroid: 4194305
};
OSF.InternalPerfMarker = {
    DataCoercionBegin: "Agave.HostCall.CoerceDataStart",
    DataCoercionEnd: "Agave.HostCall.CoerceDataEnd"
};
OSF.HostCallPerfMarker = {
    IssueCall: "Agave.HostCall.IssueCall",
    ReceiveResponse: "Agave.HostCall.ReceiveResponse",
    RuntimeExceptionRaised: "Agave.HostCall.RuntimeExecptionRaised"
};
OSF.AgaveHostAction = {
    Select: 0,
    UnSelect: 1,
    CancelDialog: 2,
    InsertAgave: 3,
    CtrlF6In: 4,
    CtrlF6Exit: 5,
    CtrlF6ExitShift: 6,
    SelectWithError: 7,
    NotifyHostError: 8,
    RefreshAddinCommands: 9,
    PageIsReady: 10,
    TabIn: 11,
    TabInShift: 12,
    TabExit: 13,
    TabExitShift: 14,
    EscExit: 15,
    F2Exit: 16,
    ExitNoFocusable: 17,
    ExitNoFocusableShift: 18,
    MouseEnter: 19,
    MouseLeave: 20,
    UpdateTargetUrl: 21,
    InstallCustomFunctions: 22,
    SendTelemetryEvent: 23,
    UninstallCustomFunctions: 24
};
OSF.SharedConstants = {NotificationConversationIdSuffix: "_ntf"};
OSF.DialogMessageType = {
    DialogMessageReceived: 0,
    DialogParentMessageReceived: 1,
    DialogClosed: 12006
};
OSF.OfficeAppContext = function OSF_OfficeAppContext(id, appName, appVersion, appUILocale, dataLocale, docUrl, clientMode, settings, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, appMinorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, clientWindowHeight, clientWindowWidth, addinName, appDomains, dialogRequirementMatrix)
{
    this._id = id;
    this._appName = appName;
    this._appVersion = appVersion;
    this._appUILocale = appUILocale;
    this._dataLocale = dataLocale;
    this._docUrl = docUrl;
    this._clientMode = clientMode;
    this._settings = settings;
    this._reason = reason;
    this._osfControlType = osfControlType;
    this._eToken = eToken;
    this._correlationId = correlationId;
    this._appInstanceId = appInstanceId;
    this._touchEnabled = touchEnabled;
    this._commerceAllowed = commerceAllowed;
    this._appMinorVersion = appMinorVersion;
    this._requirementMatrix = requirementMatrix;
    this._hostCustomMessage = hostCustomMessage;
    this._hostFullVersion = hostFullVersion;
    this._isDialog = false;
    this._clientWindowHeight = clientWindowHeight;
    this._clientWindowWidth = clientWindowWidth;
    this._addinName = addinName;
    this._appDomains = appDomains;
    this._dialogRequirementMatrix = dialogRequirementMatrix;
    this.get_id = function get_id()
    {
        return this._id
    };
    this.get_appName = function get_appName()
    {
        return this._appName
    };
    this.get_appVersion = function get_appVersion()
    {
        return this._appVersion
    };
    this.get_appUILocale = function get_appUILocale()
    {
        return this._appUILocale
    };
    this.get_dataLocale = function get_dataLocale()
    {
        return this._dataLocale
    };
    this.get_docUrl = function get_docUrl()
    {
        return this._docUrl
    };
    this.get_clientMode = function get_clientMode()
    {
        return this._clientMode
    };
    this.get_bindings = function get_bindings()
    {
        return this._bindings
    };
    this.get_settings = function get_settings()
    {
        return this._settings
    };
    this.get_reason = function get_reason()
    {
        return this._reason
    };
    this.get_osfControlType = function get_osfControlType()
    {
        return this._osfControlType
    };
    this.get_eToken = function get_eToken()
    {
        return this._eToken
    };
    this.get_correlationId = function get_correlationId()
    {
        return this._correlationId
    };
    this.get_appInstanceId = function get_appInstanceId()
    {
        return this._appInstanceId
    };
    this.get_touchEnabled = function get_touchEnabled()
    {
        return this._touchEnabled
    };
    this.get_commerceAllowed = function get_commerceAllowed()
    {
        return this._commerceAllowed
    };
    this.get_appMinorVersion = function get_appMinorVersion()
    {
        return this._appMinorVersion
    };
    this.get_requirementMatrix = function get_requirementMatrix()
    {
        return this._requirementMatrix
    };
    this.get_dialogRequirementMatrix = function get_dialogRequirementMatrix()
    {
        return this._dialogRequirementMatrix
    };
    this.get_hostCustomMessage = function get_hostCustomMessage()
    {
        return this._hostCustomMessage
    };
    this.get_hostFullVersion = function get_hostFullVersion()
    {
        return this._hostFullVersion
    };
    this.get_isDialog = function get_isDialog()
    {
        return this._isDialog
    };
    this.get_clientWindowHeight = function get_clientWindowHeight()
    {
        return this._clientWindowHeight
    };
    this.get_clientWindowWidth = function get_clientWindowWidth()
    {
        return this._clientWindowWidth
    };
    this.get_addinName = function get_addinName()
    {
        return this._addinName
    };
    this.get_appDomains = function get_appDomains()
    {
        return this._appDomains
    }
};
OSF.OsfControlType = {
    DocumentLevel: 0,
    ContainerLevel: 1
};
OSF.ClientMode = {
    ReadOnly: 0,
    ReadWrite: 1
};
OSF.OUtil.setNamespace("Microsoft",window);
OSF.OUtil.setNamespace("Office",Microsoft);
OSF.OUtil.setNamespace("Client",Microsoft.Office);
OSF.OUtil.setNamespace("WebExtension",Microsoft.Office);
Microsoft.Office.WebExtension.InitializationReason = {
    Inserted: "inserted",
    DocumentOpened: "documentOpened"
};
Microsoft.Office.WebExtension.ValueFormat = {
    Unformatted: "unformatted",
    Formatted: "formatted"
};
Microsoft.Office.WebExtension.FilterType = {All: "all"};
Microsoft.Office.WebExtension.Parameters = {
    BindingType: "bindingType",
    CoercionType: "coercionType",
    ValueFormat: "valueFormat",
    FilterType: "filterType",
    Columns: "columns",
    SampleData: "sampleData",
    GoToType: "goToType",
    SelectionMode: "selectionMode",
    Id: "id",
    PromptText: "promptText",
    ItemName: "itemName",
    FailOnCollision: "failOnCollision",
    StartRow: "startRow",
    StartColumn: "startColumn",
    RowCount: "rowCount",
    ColumnCount: "columnCount",
    Callback: "callback",
    AsyncContext: "asyncContext",
    Data: "data",
    Rows: "rows",
    OverwriteIfStale: "overwriteIfStale",
    FileType: "fileType",
    EventType: "eventType",
    Handler: "handler",
    SliceSize: "sliceSize",
    SliceIndex: "sliceIndex",
    ActiveView: "activeView",
    Status: "status",
    PlatformType: "platformType",
    HostType: "hostType",
    ForceConsent: "forceConsent",
    ForceAddAccount: "forceAddAccount",
    AuthChallenge: "authChallenge",
    Reserved: "reserved",
    Tcid: "tcid",
    Xml: "xml",
    Namespace: "namespace",
    Prefix: "prefix",
    XPath: "xPath",
    Text: "text",
    ImageLeft: "imageLeft",
    ImageTop: "imageTop",
    ImageWidth: "imageWidth",
    ImageHeight: "imageHeight",
    TaskId: "taskId",
    FieldId: "fieldId",
    FieldValue: "fieldValue",
    ServerUrl: "serverUrl",
    ListName: "listName",
    ResourceId: "resourceId",
    ViewType: "viewType",
    ViewName: "viewName",
    GetRawValue: "getRawValue",
    CellFormat: "cellFormat",
    TableOptions: "tableOptions",
    TaskIndex: "taskIndex",
    ResourceIndex: "resourceIndex",
    CustomFieldId: "customFieldId",
    Url: "url",
    MessageHandler: "messageHandler",
    Width: "width",
    Height: "height",
    RequireHTTPs: "requireHTTPS",
    MessageToParent: "messageToParent",
    DisplayInIframe: "displayInIframe",
    MessageContent: "messageContent",
    HideTitle: "hideTitle",
    UseDeviceIndependentPixels: "useDeviceIndependentPixels",
    PromptBeforeOpen: "promptBeforeOpen",
    EnforceAppDomain: "enforceAppDomain",
    AppCommandInvocationCompletedData: "appCommandInvocationCompletedData",
    Base64: "base64",
    FormId: "formId"
};
OSF.OUtil.setNamespace("DDA",OSF);
OSF.DDA.DocumentMode = {
    ReadOnly: 1,
    ReadWrite: 0
};
OSF.DDA.PropertyDescriptors = {AsyncResultStatus: "AsyncResultStatus"};
OSF.DDA.EventDescriptors = {};
OSF.DDA.ListDescriptors = {};
OSF.DDA.UI = {};
OSF.DDA.getXdmEventName = function OSF_DDA$GetXdmEventName(id, eventType)
{
    if(eventType == Microsoft.Office.WebExtension.EventType.BindingSelectionChanged || eventType == Microsoft.Office.WebExtension.EventType.BindingDataChanged || eventType == Microsoft.Office.WebExtension.EventType.DataNodeDeleted || eventType == Microsoft.Office.WebExtension.EventType.DataNodeInserted || eventType == Microsoft.Office.WebExtension.EventType.DataNodeReplaced)
        return id + "_" + eventType;
    else
        return eventType
};
OSF.DDA.MethodDispId = {
    dispidMethodMin: 64,
    dispidGetSelectedDataMethod: 64,
    dispidSetSelectedDataMethod: 65,
    dispidAddBindingFromSelectionMethod: 66,
    dispidAddBindingFromPromptMethod: 67,
    dispidGetBindingMethod: 68,
    dispidReleaseBindingMethod: 69,
    dispidGetBindingDataMethod: 70,
    dispidSetBindingDataMethod: 71,
    dispidAddRowsMethod: 72,
    dispidClearAllRowsMethod: 73,
    dispidGetAllBindingsMethod: 74,
    dispidLoadSettingsMethod: 75,
    dispidSaveSettingsMethod: 76,
    dispidGetDocumentCopyMethod: 77,
    dispidAddBindingFromNamedItemMethod: 78,
    dispidAddColumnsMethod: 79,
    dispidGetDocumentCopyChunkMethod: 80,
    dispidReleaseDocumentCopyMethod: 81,
    dispidNavigateToMethod: 82,
    dispidGetActiveViewMethod: 83,
    dispidGetDocumentThemeMethod: 84,
    dispidGetOfficeThemeMethod: 85,
    dispidGetFilePropertiesMethod: 86,
    dispidClearFormatsMethod: 87,
    dispidSetTableOptionsMethod: 88,
    dispidSetFormatsMethod: 89,
    dispidExecuteRichApiRequestMethod: 93,
    dispidAppCommandInvocationCompletedMethod: 94,
    dispidCloseContainerMethod: 97,
    dispidGetAccessTokenMethod: 98,
    dispidGetAuthContextMethod: 99,
    dispidOpenBrowserWindow: 102,
    dispidCreateDocumentMethod: 105,
    dispidInsertFormMethod: 106,
    dispidDisplayRibbonCalloutAsyncMethod: 109,
    dispidGetSelectedTaskMethod: 110,
    dispidGetSelectedResourceMethod: 111,
    dispidGetTaskMethod: 112,
    dispidGetResourceFieldMethod: 113,
    dispidGetWSSUrlMethod: 114,
    dispidGetTaskFieldMethod: 115,
    dispidGetProjectFieldMethod: 116,
    dispidGetSelectedViewMethod: 117,
    dispidGetTaskByIndexMethod: 118,
    dispidGetResourceByIndexMethod: 119,
    dispidSetTaskFieldMethod: 120,
    dispidSetResourceFieldMethod: 121,
    dispidGetMaxTaskIndexMethod: 122,
    dispidGetMaxResourceIndexMethod: 123,
    dispidCreateTaskMethod: 124,
    dispidAddDataPartMethod: 128,
    dispidGetDataPartByIdMethod: 129,
    dispidGetDataPartsByNamespaceMethod: 130,
    dispidGetDataPartXmlMethod: 131,
    dispidGetDataPartNodesMethod: 132,
    dispidDeleteDataPartMethod: 133,
    dispidGetDataNodeValueMethod: 134,
    dispidGetDataNodeXmlMethod: 135,
    dispidGetDataNodesMethod: 136,
    dispidSetDataNodeValueMethod: 137,
    dispidSetDataNodeXmlMethod: 138,
    dispidAddDataNamespaceMethod: 139,
    dispidGetDataUriByPrefixMethod: 140,
    dispidGetDataPrefixByUriMethod: 141,
    dispidGetDataNodeTextMethod: 142,
    dispidSetDataNodeTextMethod: 143,
    dispidMessageParentMethod: 144,
    dispidSendMessageMethod: 145,
    dispidExecuteFeature: 146,
    dispidQueryFeature: 147,
    dispidMethodMax: 147
};
OSF.DDA.EventDispId = {
    dispidEventMin: 0,
    dispidInitializeEvent: 0,
    dispidSettingsChangedEvent: 1,
    dispidDocumentSelectionChangedEvent: 2,
    dispidBindingSelectionChangedEvent: 3,
    dispidBindingDataChangedEvent: 4,
    dispidDocumentOpenEvent: 5,
    dispidDocumentCloseEvent: 6,
    dispidActiveViewChangedEvent: 7,
    dispidDocumentThemeChangedEvent: 8,
    dispidOfficeThemeChangedEvent: 9,
    dispidDialogMessageReceivedEvent: 10,
    dispidDialogNotificationShownInAddinEvent: 11,
    dispidDialogParentMessageReceivedEvent: 12,
    dispidObjectDeletedEvent: 13,
    dispidObjectSelectionChangedEvent: 14,
    dispidObjectDataChangedEvent: 15,
    dispidContentControlAddedEvent: 16,
    dispidActivationStatusChangedEvent: 32,
    dispidRichApiMessageEvent: 33,
    dispidAppCommandInvokedEvent: 39,
    dispidOlkItemSelectedChangedEvent: 46,
    dispidOlkRecipientsChangedEvent: 47,
    dispidOlkAppointmentTimeChangedEvent: 48,
    dispidOlkRecurrenceChangedEvent: 49,
    dispidOlkAttachmentsChangedEvent: 50,
    dispidOlkEnhancedLocationsChangedEvent: 51,
    dispidOlkInfobarClickedEvent: 52,
    dispidTaskSelectionChangedEvent: 56,
    dispidResourceSelectionChangedEvent: 57,
    dispidViewSelectionChangedEvent: 58,
    dispidDataNodeAddedEvent: 60,
    dispidDataNodeReplacedEvent: 61,
    dispidDataNodeDeletedEvent: 62,
    dispidEventMax: 63
};
OSF.DDA.ErrorCodeManager = function()
{
    var _errorMappings = {};
    return{
            getErrorArgs: function OSF_DDA_ErrorCodeManager$getErrorArgs(errorCode)
            {
                var errorArgs = _errorMappings[errorCode];
                if(!errorArgs)
                    errorArgs = _errorMappings[this.errorCodes.ooeInternalError];
                else
                {
                    if(!errorArgs.name)
                        errorArgs.name = _errorMappings[this.errorCodes.ooeInternalError].name;
                    if(!errorArgs.message)
                        errorArgs.message = _errorMappings[this.errorCodes.ooeInternalError].message
                }
                return errorArgs
            },
            addErrorMessage: function OSF_DDA_ErrorCodeManager$addErrorMessage(errorCode, errorNameMessage)
            {
                _errorMappings[errorCode] = errorNameMessage
            },
            errorCodes: {
                ooeSuccess: 0,
                ooeChunkResult: 1,
                ooeCoercionTypeNotSupported: 1e3,
                ooeGetSelectionNotMatchDataType: 1001,
                ooeCoercionTypeNotMatchBinding: 1002,
                ooeInvalidGetRowColumnCounts: 1003,
                ooeSelectionNotSupportCoercionType: 1004,
                ooeInvalidGetStartRowColumn: 1005,
                ooeNonUniformPartialGetNotSupported: 1006,
                ooeGetDataIsTooLarge: 1008,
                ooeFileTypeNotSupported: 1009,
                ooeGetDataParametersConflict: 1010,
                ooeInvalidGetColumns: 1011,
                ooeInvalidGetRows: 1012,
                ooeInvalidReadForBlankRow: 1013,
                ooeUnsupportedDataObject: 2e3,
                ooeCannotWriteToSelection: 2001,
                ooeDataNotMatchSelection: 2002,
                ooeOverwriteWorksheetData: 2003,
                ooeDataNotMatchBindingSize: 2004,
                ooeInvalidSetStartRowColumn: 2005,
                ooeInvalidDataFormat: 2006,
                ooeDataNotMatchCoercionType: 2007,
                ooeDataNotMatchBindingType: 2008,
                ooeSetDataIsTooLarge: 2009,
                ooeNonUniformPartialSetNotSupported: 2010,
                ooeInvalidSetColumns: 2011,
                ooeInvalidSetRows: 2012,
                ooeSetDataParametersConflict: 2013,
                ooeCellDataAmountBeyondLimits: 2014,
                ooeSelectionCannotBound: 3e3,
                ooeBindingNotExist: 3002,
                ooeBindingToMultipleSelection: 3003,
                ooeInvalidSelectionForBindingType: 3004,
                ooeOperationNotSupportedOnThisBindingType: 3005,
                ooeNamedItemNotFound: 3006,
                ooeMultipleNamedItemFound: 3007,
                ooeInvalidNamedItemForBindingType: 3008,
                ooeUnknownBindingType: 3009,
                ooeOperationNotSupportedOnMatrixData: 3010,
                ooeInvalidColumnsForBinding: 3011,
                ooeSettingNameNotExist: 4e3,
                ooeSettingsCannotSave: 4001,
                ooeSettingsAreStale: 4002,
                ooeOperationNotSupported: 5e3,
                ooeInternalError: 5001,
                ooeDocumentReadOnly: 5002,
                ooeEventHandlerNotExist: 5003,
                ooeInvalidApiCallInContext: 5004,
                ooeShuttingDown: 5005,
                ooeUnsupportedEnumeration: 5007,
                ooeIndexOutOfRange: 5008,
                ooeBrowserAPINotSupported: 5009,
                ooeInvalidParam: 5010,
                ooeRequestTimeout: 5011,
                ooeInvalidOrTimedOutSession: 5012,
                ooeInvalidApiArguments: 5013,
                ooeOperationCancelled: 5014,
                ooeWorkbookHidden: 5015,
                ooeTooManyIncompleteRequests: 5100,
                ooeRequestTokenUnavailable: 5101,
                ooeActivityLimitReached: 5102,
                ooeCustomXmlNodeNotFound: 6e3,
                ooeCustomXmlError: 6100,
                ooeCustomXmlExceedQuota: 6101,
                ooeCustomXmlOutOfDate: 6102,
                ooeNoCapability: 7e3,
                ooeCannotNavTo: 7001,
                ooeSpecifiedIdNotExist: 7002,
                ooeNavOutOfBound: 7004,
                ooeElementMissing: 8e3,
                ooeProtectedError: 8001,
                ooeInvalidCellsValue: 8010,
                ooeInvalidTableOptionValue: 8011,
                ooeInvalidFormatValue: 8012,
                ooeRowIndexOutOfRange: 8020,
                ooeColIndexOutOfRange: 8021,
                ooeFormatValueOutOfRange: 8022,
                ooeCellFormatAmountBeyondLimits: 8023,
                ooeMemoryFileLimit: 11e3,
                ooeNetworkProblemRetrieveFile: 11001,
                ooeInvalidSliceSize: 11002,
                ooeInvalidCallback: 11101,
                ooeInvalidWidth: 12e3,
                ooeInvalidHeight: 12001,
                ooeNavigationError: 12002,
                ooeInvalidScheme: 12003,
                ooeAppDomains: 12004,
                ooeRequireHTTPS: 12005,
                ooeWebDialogClosed: 12006,
                ooeDialogAlreadyOpened: 12007,
                ooeEndUserAllow: 12008,
                ooeEndUserIgnore: 12009,
                ooeNotUILessDialog: 12010,
                ooeCrossZone: 12011,
                ooeNotSSOAgave: 13e3,
                ooeSSOUserNotSignedIn: 13001,
                ooeSSOUserAborted: 13002,
                ooeSSOUnsupportedUserIdentity: 13003,
                ooeSSOInvalidResourceUrl: 13004,
                ooeSSOInvalidGrant: 13005,
                ooeSSOClientError: 13006,
                ooeSSOServerError: 13007,
                ooeAddinIsAlreadyRequestingToken: 13008,
                ooeSSOUserConsentNotSupportedByCurrentAddinCategory: 13009,
                ooeSSOConnectionLost: 13010,
                ooeResourceNotAllowed: 13011,
                ooeSSOUnsupportedPlatform: 13012,
                ooeAccessDenied: 13990,
                ooeGeneralException: 13991
            },
            initializeErrorMessages: function OSF_DDA_ErrorCodeManager$initializeErrorMessages(stringNS)
            {
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotSupported] = {
                    name: stringNS.L_InvalidCoercion,
                    message: stringNS.L_CoercionTypeNotSupported
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetSelectionNotMatchDataType] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_GetSelectionNotSupported
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding] = {
                    name: stringNS.L_InvalidCoercion,
                    message: stringNS.L_CoercionTypeNotMatchBinding
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRowColumnCounts] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_InvalidGetRowColumnCounts
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionNotSupportCoercionType] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_SelectionNotSupportCoercionType
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetStartRowColumn] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_InvalidGetStartRowColumn
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialGetNotSupported] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_NonUniformPartialGetNotSupported
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataIsTooLarge] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_GetDataIsTooLarge
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeFileTypeNotSupported] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_FileTypeNotSupported
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataParametersConflict] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_GetDataParametersConflict
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetColumns] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_InvalidGetColumns
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRows] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_InvalidGetRows
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidReadForBlankRow] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_InvalidReadForBlankRow
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_UnsupportedDataObject
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotWriteToSelection] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_CannotWriteToSelection
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchSelection] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_DataNotMatchSelection
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOverwriteWorksheetData] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_OverwriteWorksheetData
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingSize] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_DataNotMatchBindingSize
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetStartRowColumn] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_InvalidSetStartRowColumn
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidDataFormat] = {
                    name: stringNS.L_InvalidFormat,
                    message: stringNS.L_InvalidDataFormat
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchCoercionType] = {
                    name: stringNS.L_InvalidDataObject,
                    message: stringNS.L_DataNotMatchCoercionType
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingType] = {
                    name: stringNS.L_InvalidDataObject,
                    message: stringNS.L_DataNotMatchBindingType
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataIsTooLarge] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_SetDataIsTooLarge
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialSetNotSupported] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_NonUniformPartialSetNotSupported
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetColumns] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_InvalidSetColumns
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetRows] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_InvalidSetRows
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataParametersConflict] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_SetDataParametersConflict
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionCannotBound] = {
                    name: stringNS.L_BindingCreationError,
                    message: stringNS.L_SelectionCannotBound
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingNotExist] = {
                    name: stringNS.L_InvalidBindingError,
                    message: stringNS.L_BindingNotExist
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingToMultipleSelection] = {
                    name: stringNS.L_BindingCreationError,
                    message: stringNS.L_BindingToMultipleSelection
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSelectionForBindingType] = {
                    name: stringNS.L_BindingCreationError,
                    message: stringNS.L_InvalidSelectionForBindingType
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnThisBindingType] = {
                    name: stringNS.L_InvalidBindingOperation,
                    message: stringNS.L_OperationNotSupportedOnThisBindingType
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNamedItemNotFound] = {
                    name: stringNS.L_BindingCreationError,
                    message: stringNS.L_NamedItemNotFound
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeMultipleNamedItemFound] = {
                    name: stringNS.L_BindingCreationError,
                    message: stringNS.L_MultipleNamedItemFound
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidNamedItemForBindingType] = {
                    name: stringNS.L_BindingCreationError,
                    message: stringNS.L_InvalidNamedItemForBindingType
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnknownBindingType] = {
                    name: stringNS.L_InvalidBinding,
                    message: stringNS.L_UnknownBindingType
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnMatrixData] = {
                    name: stringNS.L_InvalidBindingOperation,
                    message: stringNS.L_OperationNotSupportedOnMatrixData
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidColumnsForBinding] = {
                    name: stringNS.L_InvalidBinding,
                    message: stringNS.L_InvalidColumnsForBinding
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingNameNotExist] = {
                    name: stringNS.L_ReadSettingsError,
                    message: stringNS.L_SettingNameNotExist
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsCannotSave] = {
                    name: stringNS.L_SaveSettingsError,
                    message: stringNS.L_SettingsCannotSave
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsAreStale] = {
                    name: stringNS.L_SettingsStaleError,
                    message: stringNS.L_SettingsAreStale
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupported] = {
                    name: stringNS.L_HostError,
                    message: stringNS.L_OperationNotSupported
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError] = {
                    name: stringNS.L_InternalError,
                    message: stringNS.L_InternalErrorDescription
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDocumentReadOnly] = {
                    name: stringNS.L_PermissionDenied,
                    message: stringNS.L_DocumentReadOnly
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist] = {
                    name: stringNS.L_EventRegistrationError,
                    message: stringNS.L_EventHandlerNotExist
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext] = {
                    name: stringNS.L_InvalidAPICall,
                    message: stringNS.L_InvalidApiCallInContext
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeShuttingDown] = {
                    name: stringNS.L_ShuttingDown,
                    message: stringNS.L_ShuttingDown
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration] = {
                    name: stringNS.L_UnsupportedEnumeration,
                    message: stringNS.L_UnsupportedEnumerationMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange] = {
                    name: stringNS.L_IndexOutOfRange,
                    message: stringNS.L_IndexOutOfRange
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBrowserAPINotSupported] = {
                    name: stringNS.L_APINotSupported,
                    message: stringNS.L_BrowserAPINotSupported
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTimeout] = {
                    name: stringNS.L_APICallFailed,
                    message: stringNS.L_RequestTimeout
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidOrTimedOutSession] = {
                    name: stringNS.L_InvalidOrTimedOutSession,
                    message: stringNS.L_InvalidOrTimedOutSessionMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeTooManyIncompleteRequests] = {
                    name: stringNS.L_APICallFailed,
                    message: stringNS.L_TooManyIncompleteRequests
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTokenUnavailable] = {
                    name: stringNS.L_APICallFailed,
                    message: stringNS.L_RequestTokenUnavailable
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeActivityLimitReached] = {
                    name: stringNS.L_APICallFailed,
                    message: stringNS.L_ActivityLimitReached
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiArguments] = {
                    name: stringNS.L_APICallFailed,
                    message: stringNS.L_InvalidApiArgumentsMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeWorkbookHidden] = {
                    name: stringNS.L_APICallFailed,
                    message: stringNS.L_WorkbookHiddenMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlNodeNotFound] = {
                    name: stringNS.L_InvalidNode,
                    message: stringNS.L_CustomXmlNodeNotFound
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlError] = {
                    name: stringNS.L_CustomXmlError,
                    message: stringNS.L_CustomXmlError
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlExceedQuota] = {
                    name: stringNS.L_CustomXmlExceedQuotaName,
                    message: stringNS.L_CustomXmlExceedQuotaMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlOutOfDate] = {
                    name: stringNS.L_CustomXmlOutOfDateName,
                    message: stringNS.L_CustomXmlOutOfDateMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability] = {
                    name: stringNS.L_PermissionDenied,
                    message: stringNS.L_NoCapability
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotNavTo] = {
                    name: stringNS.L_CannotNavigateTo,
                    message: stringNS.L_CannotNavigateTo
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSpecifiedIdNotExist] = {
                    name: stringNS.L_SpecifiedIdNotExist,
                    message: stringNS.L_SpecifiedIdNotExist
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavOutOfBound] = {
                    name: stringNS.L_NavOutOfBound,
                    message: stringNS.L_NavOutOfBound
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellDataAmountBeyondLimits] = {
                    name: stringNS.L_DataWriteReminder,
                    message: stringNS.L_CellDataAmountBeyondLimits
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeElementMissing] = {
                    name: stringNS.L_MissingParameter,
                    message: stringNS.L_ElementMissing
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeProtectedError] = {
                    name: stringNS.L_PermissionDenied,
                    message: stringNS.L_NoCapability
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCellsValue] = {
                    name: stringNS.L_InvalidValue,
                    message: stringNS.L_InvalidCellsValue
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidTableOptionValue] = {
                    name: stringNS.L_InvalidValue,
                    message: stringNS.L_InvalidTableOptionValue
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidFormatValue] = {
                    name: stringNS.L_InvalidValue,
                    message: stringNS.L_InvalidFormatValue
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRowIndexOutOfRange] = {
                    name: stringNS.L_OutOfRange,
                    message: stringNS.L_RowIndexOutOfRange
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeColIndexOutOfRange] = {
                    name: stringNS.L_OutOfRange,
                    message: stringNS.L_ColIndexOutOfRange
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeFormatValueOutOfRange] = {
                    name: stringNS.L_OutOfRange,
                    message: stringNS.L_FormatValueOutOfRange
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellFormatAmountBeyondLimits] = {
                    name: stringNS.L_FormattingReminder,
                    message: stringNS.L_CellFormatAmountBeyondLimits
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeMemoryFileLimit] = {
                    name: stringNS.L_MemoryLimit,
                    message: stringNS.L_CloseFileBeforeRetrieve
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNetworkProblemRetrieveFile] = {
                    name: stringNS.L_NetworkProblem,
                    message: stringNS.L_NetworkProblemRetrieveFile
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize] = {
                    name: stringNS.L_InvalidValue,
                    message: stringNS.L_SliceSizeNotSupported
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened] = {
                    name: stringNS.L_DisplayDialogError,
                    message: stringNS.L_DialogAlreadyOpened
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidWidth] = {
                    name: stringNS.L_IndexOutOfRange,
                    message: stringNS.L_IndexOutOfRange
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidHeight] = {
                    name: stringNS.L_IndexOutOfRange,
                    message: stringNS.L_IndexOutOfRange
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavigationError] = {
                    name: stringNS.L_DisplayDialogError,
                    message: stringNS.L_NetworkProblem
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidScheme] = {
                    name: stringNS.L_DialogNavigateError,
                    message: stringNS.L_DialogInvalidScheme
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeAppDomains] = {
                    name: stringNS.L_DisplayDialogError,
                    message: stringNS.L_DialogAddressNotTrusted
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequireHTTPS] = {
                    name: stringNS.L_DisplayDialogError,
                    message: stringNS.L_DialogRequireHTTPS
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeEndUserIgnore] = {
                    name: stringNS.L_DisplayDialogError,
                    message: stringNS.L_UserClickIgnore
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCrossZone] = {
                    name: stringNS.L_DisplayDialogError,
                    message: stringNS.L_NewWindowCrossZoneErrorString
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNotSSOAgave] = {
                    name: stringNS.L_APINotSupported,
                    message: stringNS.L_InvalidSSOAddinMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserNotSignedIn] = {
                    name: stringNS.L_UserNotSignedIn,
                    message: stringNS.L_UserNotSignedIn
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserAborted] = {
                    name: stringNS.L_UserAborted,
                    message: stringNS.L_UserAbortedMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUnsupportedUserIdentity] = {
                    name: stringNS.L_UnsupportedUserIdentity,
                    message: stringNS.L_UnsupportedUserIdentityMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidResourceUrl] = {
                    name: stringNS.L_InvalidResourceUrl,
                    message: stringNS.L_InvalidResourceUrlMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidGrant] = {
                    name: stringNS.L_InvalidGrant,
                    message: stringNS.L_InvalidGrantMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOClientError] = {
                    name: stringNS.L_SSOClientError,
                    message: stringNS.L_SSOClientErrorMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOServerError] = {
                    name: stringNS.L_SSOServerError,
                    message: stringNS.L_SSOServerErrorMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeAddinIsAlreadyRequestingToken] = {
                    name: stringNS.L_AddinIsAlreadyRequestingToken,
                    message: stringNS.L_AddinIsAlreadyRequestingTokenMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserConsentNotSupportedByCurrentAddinCategory] = {
                    name: stringNS.L_SSOUserConsentNotSupportedByCurrentAddinCategory,
                    message: stringNS.L_SSOUserConsentNotSupportedByCurrentAddinCategoryMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOConnectionLost] = {
                    name: stringNS.L_SSOConnectionLostError,
                    message: stringNS.L_SSOConnectionLostErrorMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUnsupportedPlatform] = {
                    name: stringNS.L_SSOConnectionLostError,
                    message: stringNS.L_SSOUnsupportedPlatform
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationCancelled] = {
                    name: stringNS.L_OperationCancelledError,
                    message: stringNS.L_OperationCancelledErrorMessage
                }
            }
        }
}();
var OfficeExt;
(function(OfficeExt)
{
    var Requirement;
    (function(Requirement)
    {
        var RequirementVersion = function()
            {
                function RequirementVersion(){}
                return RequirementVersion
            }();
        Requirement.RequirementVersion = RequirementVersion;
        var RequirementMatrix = function()
            {
                function RequirementMatrix(_setMap)
                {
                    this.isSetSupported = function _isSetSupported(name, minVersion)
                    {
                        if(name == undefined)
                            return false;
                        if(minVersion == undefined)
                            minVersion = 0;
                        var setSupportArray = this._setMap;
                        var sets = setSupportArray._sets;
                        if(sets.hasOwnProperty(name.toLowerCase()))
                        {
                            var setMaxVersion = sets[name.toLowerCase()];
                            try
                            {
                                var setMaxVersionNum = this._getVersion(setMaxVersion);
                                minVersion = minVersion + "";
                                var minVersionNum = this._getVersion(minVersion);
                                if(setMaxVersionNum.major > 0 && setMaxVersionNum.major > minVersionNum.major)
                                    return true;
                                if(setMaxVersionNum.minor > 0 && setMaxVersionNum.minor > 0 && setMaxVersionNum.major == minVersionNum.major && setMaxVersionNum.minor >= minVersionNum.minor)
                                    return true
                            }
                            catch(e)
                            {
                                return false
                            }
                        }
                        return false
                    };
                    this._getVersion = function(version)
                    {
                        version = version + "";
                        var temp = version.split(".");
                        var major = 0;
                        var minor = 0;
                        if(temp.length < 2 && isNaN(Number(version)))
                            throw"version format incorrect";
                        else
                        {
                            major = Number(temp[0]);
                            if(temp.length >= 2)
                                minor = Number(temp[1]);
                            if(isNaN(major) || isNaN(minor))
                                throw"version format incorrect";
                        }
                        var result = {
                                minor: minor,
                                major: major
                            };
                        return result
                    };
                    this._setMap = _setMap;
                    this.isSetSupported = this.isSetSupported.bind(this)
                }
                return RequirementMatrix
            }();
        Requirement.RequirementMatrix = RequirementMatrix;
        var DefaultSetRequirement = function()
            {
                function DefaultSetRequirement(setMap)
                {
                    this._addSetMap = function DefaultSetRequirement_addSetMap(addedSet)
                    {
                        for(var name in addedSet)
                            this._sets[name] = addedSet[name]
                    };
                    this._sets = setMap
                }
                return DefaultSetRequirement
            }();
        Requirement.DefaultSetRequirement = DefaultSetRequirement;
        var DefaultDialogSetRequirement = function(_super)
            {
                __extends(DefaultDialogSetRequirement,_super);
                function DefaultDialogSetRequirement()
                {
                    _super.call(this,{dialogapi: 1.1})
                }
                return DefaultDialogSetRequirement
            }(DefaultSetRequirement);
        Requirement.DefaultDialogSetRequirement = DefaultDialogSetRequirement;
        var ExcelClientDefaultSetRequirement = function(_super)
            {
                __extends(ExcelClientDefaultSetRequirement,_super);
                function ExcelClientDefaultSetRequirement()
                {
                    _super.call(this,{
                        bindingevents: 1.1,
                        documentevents: 1.1,
                        excelapi: 1.1,
                        matrixbindings: 1.1,
                        matrixcoercion: 1.1,
                        selection: 1.1,
                        settings: 1.1,
                        tablebindings: 1.1,
                        tablecoercion: 1.1,
                        textbindings: 1.1,
                        textcoercion: 1.1
                    })
                }
                return ExcelClientDefaultSetRequirement
            }(DefaultSetRequirement);
        Requirement.ExcelClientDefaultSetRequirement = ExcelClientDefaultSetRequirement;
        var ExcelClientV1DefaultSetRequirement = function(_super)
            {
                __extends(ExcelClientV1DefaultSetRequirement,_super);
                function ExcelClientV1DefaultSetRequirement()
                {
                    _super.call(this);
                    this._addSetMap({imagecoercion: 1.1})
                }
                return ExcelClientV1DefaultSetRequirement
            }(ExcelClientDefaultSetRequirement);
        Requirement.ExcelClientV1DefaultSetRequirement = ExcelClientV1DefaultSetRequirement;
        var OutlookClientDefaultSetRequirement = function(_super)
            {
                __extends(OutlookClientDefaultSetRequirement,_super);
                function OutlookClientDefaultSetRequirement()
                {
                    _super.call(this,{mailbox: 1.3})
                }
                return OutlookClientDefaultSetRequirement
            }(DefaultSetRequirement);
        Requirement.OutlookClientDefaultSetRequirement = OutlookClientDefaultSetRequirement;
        var WordClientDefaultSetRequirement = function(_super)
            {
                __extends(WordClientDefaultSetRequirement,_super);
                function WordClientDefaultSetRequirement()
                {
                    _super.call(this,{
                        bindingevents: 1.1,
                        compressedfile: 1.1,
                        customxmlparts: 1.1,
                        documentevents: 1.1,
                        file: 1.1,
                        htmlcoercion: 1.1,
                        matrixbindings: 1.1,
                        matrixcoercion: 1.1,
                        ooxmlcoercion: 1.1,
                        pdffile: 1.1,
                        selection: 1.1,
                        settings: 1.1,
                        tablebindings: 1.1,
                        tablecoercion: 1.1,
                        textbindings: 1.1,
                        textcoercion: 1.1,
                        textfile: 1.1,
                        wordapi: 1.1
                    })
                }
                return WordClientDefaultSetRequirement
            }(DefaultSetRequirement);
        Requirement.WordClientDefaultSetRequirement = WordClientDefaultSetRequirement;
        var WordClientV1DefaultSetRequirement = function(_super)
            {
                __extends(WordClientV1DefaultSetRequirement,_super);
                function WordClientV1DefaultSetRequirement()
                {
                    _super.call(this);
                    this._addSetMap({
                        customxmlparts: 1.2,
                        wordapi: 1.2,
                        imagecoercion: 1.1
                    })
                }
                return WordClientV1DefaultSetRequirement
            }(WordClientDefaultSetRequirement);
        Requirement.WordClientV1DefaultSetRequirement = WordClientV1DefaultSetRequirement;
        var PowerpointClientDefaultSetRequirement = function(_super)
            {
                __extends(PowerpointClientDefaultSetRequirement,_super);
                function PowerpointClientDefaultSetRequirement()
                {
                    _super.call(this,{
                        activeview: 1.1,
                        compressedfile: 1.1,
                        documentevents: 1.1,
                        file: 1.1,
                        pdffile: 1.1,
                        selection: 1.1,
                        settings: 1.1,
                        textcoercion: 1.1
                    })
                }
                return PowerpointClientDefaultSetRequirement
            }(DefaultSetRequirement);
        Requirement.PowerpointClientDefaultSetRequirement = PowerpointClientDefaultSetRequirement;
        var PowerpointClientV1DefaultSetRequirement = function(_super)
            {
                __extends(PowerpointClientV1DefaultSetRequirement,_super);
                function PowerpointClientV1DefaultSetRequirement()
                {
                    _super.call(this);
                    this._addSetMap({imagecoercion: 1.1})
                }
                return PowerpointClientV1DefaultSetRequirement
            }(PowerpointClientDefaultSetRequirement);
        Requirement.PowerpointClientV1DefaultSetRequirement = PowerpointClientV1DefaultSetRequirement;
        var ProjectClientDefaultSetRequirement = function(_super)
            {
                __extends(ProjectClientDefaultSetRequirement,_super);
                function ProjectClientDefaultSetRequirement()
                {
                    _super.call(this,{
                        selection: 1.1,
                        textcoercion: 1.1
                    })
                }
                return ProjectClientDefaultSetRequirement
            }(DefaultSetRequirement);
        Requirement.ProjectClientDefaultSetRequirement = ProjectClientDefaultSetRequirement;
        var ExcelWebDefaultSetRequirement = function(_super)
            {
                __extends(ExcelWebDefaultSetRequirement,_super);
                function ExcelWebDefaultSetRequirement()
                {
                    _super.call(this,{
                        bindingevents: 1.1,
                        documentevents: 1.1,
                        matrixbindings: 1.1,
                        matrixcoercion: 1.1,
                        selection: 1.1,
                        settings: 1.1,
                        tablebindings: 1.1,
                        tablecoercion: 1.1,
                        textbindings: 1.1,
                        textcoercion: 1.1,
                        file: 1.1
                    })
                }
                return ExcelWebDefaultSetRequirement
            }(DefaultSetRequirement);
        Requirement.ExcelWebDefaultSetRequirement = ExcelWebDefaultSetRequirement;
        var WordWebDefaultSetRequirement = function(_super)
            {
                __extends(WordWebDefaultSetRequirement,_super);
                function WordWebDefaultSetRequirement()
                {
                    _super.call(this,{
                        compressedfile: 1.1,
                        documentevents: 1.1,
                        file: 1.1,
                        imagecoercion: 1.1,
                        matrixcoercion: 1.1,
                        ooxmlcoercion: 1.1,
                        pdffile: 1.1,
                        selection: 1.1,
                        settings: 1.1,
                        tablecoercion: 1.1,
                        textcoercion: 1.1,
                        textfile: 1.1
                    })
                }
                return WordWebDefaultSetRequirement
            }(DefaultSetRequirement);
        Requirement.WordWebDefaultSetRequirement = WordWebDefaultSetRequirement;
        var PowerpointWebDefaultSetRequirement = function(_super)
            {
                __extends(PowerpointWebDefaultSetRequirement,_super);
                function PowerpointWebDefaultSetRequirement()
                {
                    _super.call(this,{
                        activeview: 1.1,
                        settings: 1.1
                    })
                }
                return PowerpointWebDefaultSetRequirement
            }(DefaultSetRequirement);
        Requirement.PowerpointWebDefaultSetRequirement = PowerpointWebDefaultSetRequirement;
        var OutlookWebDefaultSetRequirement = function(_super)
            {
                __extends(OutlookWebDefaultSetRequirement,_super);
                function OutlookWebDefaultSetRequirement()
                {
                    _super.call(this,{mailbox: 1.3})
                }
                return OutlookWebDefaultSetRequirement
            }(DefaultSetRequirement);
        Requirement.OutlookWebDefaultSetRequirement = OutlookWebDefaultSetRequirement;
        var SwayWebDefaultSetRequirement = function(_super)
            {
                __extends(SwayWebDefaultSetRequirement,_super);
                function SwayWebDefaultSetRequirement()
                {
                    _super.call(this,{
                        activeview: 1.1,
                        documentevents: 1.1,
                        selection: 1.1,
                        settings: 1.1,
                        textcoercion: 1.1
                    })
                }
                return SwayWebDefaultSetRequirement
            }(DefaultSetRequirement);
        Requirement.SwayWebDefaultSetRequirement = SwayWebDefaultSetRequirement;
        var AccessWebDefaultSetRequirement = function(_super)
            {
                __extends(AccessWebDefaultSetRequirement,_super);
                function AccessWebDefaultSetRequirement()
                {
                    _super.call(this,{
                        bindingevents: 1.1,
                        partialtablebindings: 1.1,
                        settings: 1.1,
                        tablebindings: 1.1,
                        tablecoercion: 1.1
                    })
                }
                return AccessWebDefaultSetRequirement
            }(DefaultSetRequirement);
        Requirement.AccessWebDefaultSetRequirement = AccessWebDefaultSetRequirement;
        var ExcelIOSDefaultSetRequirement = function(_super)
            {
                __extends(ExcelIOSDefaultSetRequirement,_super);
                function ExcelIOSDefaultSetRequirement()
                {
                    _super.call(this,{
                        bindingevents: 1.1,
                        documentevents: 1.1,
                        matrixbindings: 1.1,
                        matrixcoercion: 1.1,
                        selection: 1.1,
                        settings: 1.1,
                        tablebindings: 1.1,
                        tablecoercion: 1.1,
                        textbindings: 1.1,
                        textcoercion: 1.1
                    })
                }
                return ExcelIOSDefaultSetRequirement
            }(DefaultSetRequirement);
        Requirement.ExcelIOSDefaultSetRequirement = ExcelIOSDefaultSetRequirement;
        var WordIOSDefaultSetRequirement = function(_super)
            {
                __extends(WordIOSDefaultSetRequirement,_super);
                function WordIOSDefaultSetRequirement()
                {
                    _super.call(this,{
                        bindingevents: 1.1,
                        compressedfile: 1.1,
                        customxmlparts: 1.1,
                        documentevents: 1.1,
                        file: 1.1,
                        htmlcoercion: 1.1,
                        matrixbindings: 1.1,
                        matrixcoercion: 1.1,
                        ooxmlcoercion: 1.1,
                        pdffile: 1.1,
                        selection: 1.1,
                        settings: 1.1,
                        tablebindings: 1.1,
                        tablecoercion: 1.1,
                        textbindings: 1.1,
                        textcoercion: 1.1,
                        textfile: 1.1
                    })
                }
                return WordIOSDefaultSetRequirement
            }(DefaultSetRequirement);
        Requirement.WordIOSDefaultSetRequirement = WordIOSDefaultSetRequirement;
        var WordIOSV1DefaultSetRequirement = function(_super)
            {
                __extends(WordIOSV1DefaultSetRequirement,_super);
                function WordIOSV1DefaultSetRequirement()
                {
                    _super.call(this);
                    this._addSetMap({
                        customxmlparts: 1.2,
                        wordapi: 1.2
                    })
                }
                return WordIOSV1DefaultSetRequirement
            }(WordIOSDefaultSetRequirement);
        Requirement.WordIOSV1DefaultSetRequirement = WordIOSV1DefaultSetRequirement;
        var PowerpointIOSDefaultSetRequirement = function(_super)
            {
                __extends(PowerpointIOSDefaultSetRequirement,_super);
                function PowerpointIOSDefaultSetRequirement()
                {
                    _super.call(this,{
                        activeview: 1.1,
                        compressedfile: 1.1,
                        documentevents: 1.1,
                        file: 1.1,
                        pdffile: 1.1,
                        selection: 1.1,
                        settings: 1.1,
                        textcoercion: 1.1
                    })
                }
                return PowerpointIOSDefaultSetRequirement
            }(DefaultSetRequirement);
        Requirement.PowerpointIOSDefaultSetRequirement = PowerpointIOSDefaultSetRequirement;
        var OutlookIOSDefaultSetRequirement = function(_super)
            {
                __extends(OutlookIOSDefaultSetRequirement,_super);
                function OutlookIOSDefaultSetRequirement()
                {
                    _super.call(this,{mailbox: 1.1})
                }
                return OutlookIOSDefaultSetRequirement
            }(DefaultSetRequirement);
        Requirement.OutlookIOSDefaultSetRequirement = OutlookIOSDefaultSetRequirement;
        var RequirementsMatrixFactory = function()
            {
                function RequirementsMatrixFactory(){}
                RequirementsMatrixFactory.initializeOsfDda = function()
                {
                    OSF.OUtil.setNamespace("Requirement",OSF.DDA)
                };
                RequirementsMatrixFactory.getDefaultRequirementMatrix = function(appContext)
                {
                    this.initializeDefaultSetMatrix();
                    var defaultRequirementMatrix = undefined;
                    var clientRequirement = appContext.get_requirementMatrix();
                    if(clientRequirement != undefined && clientRequirement.length > 0 && typeof JSON !== "undefined")
                    {
                        var matrixItem = JSON.parse(appContext.get_requirementMatrix().toLowerCase());
                        defaultRequirementMatrix = new RequirementMatrix(new DefaultSetRequirement(matrixItem))
                    }
                    else
                    {
                        var appLocator = RequirementsMatrixFactory.getClientFullVersionString(appContext);
                        if(RequirementsMatrixFactory.DefaultSetArrayMatrix != undefined && RequirementsMatrixFactory.DefaultSetArrayMatrix[appLocator] != undefined)
                            defaultRequirementMatrix = new RequirementMatrix(RequirementsMatrixFactory.DefaultSetArrayMatrix[appLocator]);
                        else
                            defaultRequirementMatrix = new RequirementMatrix(new DefaultSetRequirement({}))
                    }
                    return defaultRequirementMatrix
                };
                RequirementsMatrixFactory.getDefaultDialogRequirementMatrix = function(appContext)
                {
                    var defaultRequirementMatrix = undefined;
                    var clientRequirement = appContext.get_dialogRequirementMatrix();
                    if(clientRequirement != undefined && clientRequirement.length > 0 && typeof JSON !== "undefined")
                    {
                        var matrixItem = JSON.parse(appContext.get_requirementMatrix().toLowerCase());
                        defaultRequirementMatrix = new RequirementMatrix(new DefaultSetRequirement(matrixItem))
                    }
                    else
                        defaultRequirementMatrix = new RequirementMatrix(new DefaultDialogSetRequirement);
                    return defaultRequirementMatrix
                };
                RequirementsMatrixFactory.getClientFullVersionString = function(appContext)
                {
                    var appMinorVersion = appContext.get_appMinorVersion();
                    var appMinorVersionString = "";
                    var appFullVersion = "";
                    var appName = appContext.get_appName();
                    var isIOSClient = appName == 1024 || appName == 4096 || appName == 8192 || appName == 65536;
                    if(isIOSClient && appContext.get_appVersion() == 1)
                        if(appName == 4096 && appMinorVersion >= 15)
                            appFullVersion = "16.00.01";
                        else
                            appFullVersion = "16.00";
                    else if(appContext.get_appName() == 64)
                        appFullVersion = appContext.get_appVersion();
                    else
                    {
                        if(appMinorVersion < 10)
                            appMinorVersionString = "0" + appMinorVersion;
                        else
                            appMinorVersionString = "" + appMinorVersion;
                        appFullVersion = appContext.get_appVersion() + "." + appMinorVersionString
                    }
                    return appContext.get_appName() + "-" + appFullVersion
                };
                RequirementsMatrixFactory.initializeDefaultSetMatrix = function()
                {
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_RCLIENT_1600] = new ExcelClientDefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_RCLIENT_1600] = new WordClientDefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_RCLIENT_1600] = new PowerpointClientDefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_RCLIENT_1601] = new ExcelClientV1DefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_RCLIENT_1601] = new WordClientV1DefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_RCLIENT_1601] = new PowerpointClientV1DefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_RCLIENT_1600] = new OutlookClientDefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_WAC_1600] = new ExcelWebDefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_WAC_1600] = new WordWebDefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_WAC_1600] = new OutlookWebDefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_WAC_1601] = new OutlookWebDefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Project_RCLIENT_1600] = new ProjectClientDefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Access_WAC_1600] = new AccessWebDefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_WAC_1600] = new PowerpointWebDefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_IOS_1600] = new ExcelIOSDefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.SWAY_WAC_1600] = new SwayWebDefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_IOS_1600] = new WordIOSDefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_IOS_16001] = new WordIOSV1DefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_IOS_1600] = new PowerpointIOSDefaultSetRequirement;
                    RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_IOS_1600] = new OutlookIOSDefaultSetRequirement
                };
                RequirementsMatrixFactory.Excel_RCLIENT_1600 = "1-16.00";
                RequirementsMatrixFactory.Excel_RCLIENT_1601 = "1-16.01";
                RequirementsMatrixFactory.Word_RCLIENT_1600 = "2-16.00";
                RequirementsMatrixFactory.Word_RCLIENT_1601 = "2-16.01";
                RequirementsMatrixFactory.PowerPoint_RCLIENT_1600 = "4-16.00";
                RequirementsMatrixFactory.PowerPoint_RCLIENT_1601 = "4-16.01";
                RequirementsMatrixFactory.Outlook_RCLIENT_1600 = "8-16.00";
                RequirementsMatrixFactory.Excel_WAC_1600 = "16-16.00";
                RequirementsMatrixFactory.Word_WAC_1600 = "32-16.00";
                RequirementsMatrixFactory.Outlook_WAC_1600 = "64-16.00";
                RequirementsMatrixFactory.Outlook_WAC_1601 = "64-16.01";
                RequirementsMatrixFactory.Project_RCLIENT_1600 = "128-16.00";
                RequirementsMatrixFactory.Access_WAC_1600 = "256-16.00";
                RequirementsMatrixFactory.PowerPoint_WAC_1600 = "512-16.00";
                RequirementsMatrixFactory.Excel_IOS_1600 = "1024-16.00";
                RequirementsMatrixFactory.SWAY_WAC_1600 = "2048-16.00";
                RequirementsMatrixFactory.Word_IOS_1600 = "4096-16.00";
                RequirementsMatrixFactory.Word_IOS_16001 = "4096-16.00.01";
                RequirementsMatrixFactory.PowerPoint_IOS_1600 = "8192-16.00";
                RequirementsMatrixFactory.Outlook_IOS_1600 = "65536-16.00";
                RequirementsMatrixFactory.DefaultSetArrayMatrix = {};
                return RequirementsMatrixFactory
            }();
        Requirement.RequirementsMatrixFactory = RequirementsMatrixFactory
    })(Requirement = OfficeExt.Requirement || (OfficeExt.Requirement = {}))
})(OfficeExt || (OfficeExt = {}));
OfficeExt.Requirement.RequirementsMatrixFactory.initializeOsfDda();
Microsoft.Office.WebExtension.ApplicationMode = {
    WebEditor: "webEditor",
    WebViewer: "webViewer",
    Client: "client"
};
Microsoft.Office.WebExtension.DocumentMode = {
    ReadOnly: "readOnly",
    ReadWrite: "readWrite"
};
OSF.NamespaceManager = function OSF_NamespaceManager()
{
    var _userOffice;
    var _useShortcut = false;
    return{
            enableShortcut: function OSF_NamespaceManager$enableShortcut()
            {
                if(!_useShortcut)
                {
                    if(window.Office)
                        _userOffice = window.Office;
                    else
                        OSF.OUtil.setNamespace("Office",window);
                    window.Office = Microsoft.Office.WebExtension;
                    _useShortcut = true
                }
            },
            disableShortcut: function OSF_NamespaceManager$disableShortcut()
            {
                if(_useShortcut)
                {
                    if(_userOffice)
                        window.Office = _userOffice;
                    else
                        OSF.OUtil.unsetNamespace("Office",window);
                    _useShortcut = false
                }
            }
        }
}();
OSF.NamespaceManager.enableShortcut();
Microsoft.Office.WebExtension.useShortNamespace = function Microsoft_Office_WebExtension_useShortNamespace(useShortcut)
{
    if(useShortcut)
        OSF.NamespaceManager.enableShortcut();
    else
        OSF.NamespaceManager.disableShortcut()
};
Microsoft.Office.WebExtension.select = function Microsoft_Office_WebExtension_select(str, errorCallback)
{
    var promise;
    if(str && typeof str == "string")
    {
        var index = str.indexOf("#");
        if(index != -1)
        {
            var op = str.substring(0,index);
            var target = str.substring(index + 1);
            switch(op)
            {
                case"binding":
                case"bindings":
                    if(target)
                        promise = new OSF.DDA.BindingPromise(target);
                    break
            }
        }
    }
    if(!promise)
    {
        if(errorCallback)
        {
            var callbackType = typeof errorCallback;
            if(callbackType == "function")
            {
                var callArgs = {};
                callArgs[Microsoft.Office.WebExtension.Parameters.Callback] = errorCallback;
                OSF.DDA.issueAsyncResult(callArgs,OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext,OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext))
            }
            else
                throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction,callbackType);
        }
    }
    else
    {
        promise.onFail = errorCallback;
        return promise
    }
};
OSF.DDA.Context = function OSF_DDA_Context(officeAppContext, document, license, appOM, getOfficeTheme)
{
    OSF.OUtil.defineEnumerableProperties(this,{
        contentLanguage: {value: officeAppContext.get_dataLocale()},
        displayLanguage: {value: officeAppContext.get_appUILocale()},
        touchEnabled: {value: officeAppContext.get_touchEnabled()},
        commerceAllowed: {value: officeAppContext.get_commerceAllowed()},
        host: {value: OfficeExt.HostName.Host.getInstance().getHost()},
        platform: {value: OfficeExt.HostName.Host.getInstance().getPlatform()},
        isDialog: {value: OSF._OfficeAppFactory.getHostInfo().isDialog},
        diagnostics: {value: OfficeExt.HostName.Host.getInstance().getDiagnostics(officeAppContext.get_hostFullVersion())}
    });
    if(license)
        OSF.OUtil.defineEnumerableProperty(this,"license",{value: license});
    if(officeAppContext.ui)
        OSF.OUtil.defineEnumerableProperty(this,"ui",{value: officeAppContext.ui});
    if(officeAppContext.auth)
        OSF.OUtil.defineEnumerableProperty(this,"auth",{value: officeAppContext.auth});
    if(officeAppContext.webAuth)
        OSF.OUtil.defineEnumerableProperty(this,"webAuth",{value: officeAppContext.webAuth});
    if(officeAppContext.application)
        OSF.OUtil.defineEnumerableProperty(this,"application",{value: officeAppContext.application});
    if(officeAppContext.get_isDialog())
    {
        var requirements = OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultDialogRequirementMatrix(officeAppContext);
        OSF.OUtil.defineEnumerableProperty(this,"requirements",{value: requirements})
    }
    else
    {
        if(document)
            OSF.OUtil.defineEnumerableProperty(this,"document",{value: document});
        if(appOM)
        {
            var displayName = appOM.displayName || "appOM";
            delete appOM.displayName;
            OSF.OUtil.defineEnumerableProperty(this,displayName,{value: appOM})
        }
        if(getOfficeTheme)
            OSF.OUtil.defineEnumerableProperty(this,"officeTheme",{get: function()
                {
                    return getOfficeTheme()
                }});
        var requirements = OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultRequirementMatrix(officeAppContext);
        OSF.OUtil.defineEnumerableProperty(this,"requirements",{value: requirements})
    }
};
OSF.DDA.OutlookContext = function OSF_DDA_OutlookContext(appContext, settings, license, appOM, getOfficeTheme)
{
    OSF.DDA.OutlookContext.uber.constructor.call(this,appContext,null,license,appOM,getOfficeTheme);
    if(settings)
        OSF.OUtil.defineEnumerableProperty(this,"roamingSettings",{value: settings})
};
OSF.OUtil.extend(OSF.DDA.OutlookContext,OSF.DDA.Context);
OSF.DDA.OutlookAppOm = function OSF_DDA_OutlookAppOm(appContext, window, appReady){};
OSF.DDA.Application = function OSF_DDA_Application(officeAppContext){};
OSF.DDA.Document = function OSF_DDA_Document(officeAppContext, settings)
{
    var mode;
    switch(officeAppContext.get_clientMode())
    {
        case OSF.ClientMode.ReadOnly:
            mode = Microsoft.Office.WebExtension.DocumentMode.ReadOnly;
            break;
        case OSF.ClientMode.ReadWrite:
            mode = Microsoft.Office.WebExtension.DocumentMode.ReadWrite;
            break
    }
    if(settings)
        OSF.OUtil.defineEnumerableProperty(this,"settings",{value: settings});
    OSF.OUtil.defineMutableProperties(this,{
        mode: {value: mode},
        url: {value: officeAppContext.get_docUrl()}
    })
};
OSF.DDA.JsomDocument = function OSF_DDA_JsomDocument(officeAppContext, bindingFacade, settings)
{
    OSF.DDA.JsomDocument.uber.constructor.call(this,officeAppContext,settings);
    if(bindingFacade)
        OSF.OUtil.defineEnumerableProperty(this,"bindings",{get: function OSF_DDA_Document$GetBindings()
            {
                return bindingFacade
            }});
    var am = OSF.DDA.AsyncMethodNames;
    OSF.DDA.DispIdHost.addAsyncMethods(this,[am.GetSelectedDataAsync,am.SetSelectedDataAsync]);
    OSF.DDA.DispIdHost.addEventSupport(this,new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged]))
};
OSF.OUtil.extend(OSF.DDA.JsomDocument,OSF.DDA.Document);
OSF.OUtil.defineEnumerableProperty(Microsoft.Office.WebExtension,"context",{get: function Microsoft_Office_WebExtension$GetContext()
    {
        var context;
        if(OSF && OSF._OfficeAppFactory)
            context = OSF._OfficeAppFactory.getContext();
        return context
    }});
OSF.DDA.License = function OSF_DDA_License(eToken)
{
    OSF.OUtil.defineEnumerableProperty(this,"value",{value: eToken})
};
OSF.DDA.ApiMethodCall = function OSF_DDA_ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName)
{
    var requiredCount = requiredParameters.length;
    var getInvalidParameterString = OSF.OUtil.delayExecutionAndCache(function()
        {
            return OSF.OUtil.formatString(Strings.OfficeOM.L_InvalidParameters,displayName)
        });
    this.verifyArguments = function OSF_DDA_ApiMethodCall$VerifyArguments(params, args)
    {
        for(var name in params)
        {
            var param = params[name];
            var arg = args[name];
            if(param["enum"])
                switch(typeof arg)
                {
                    case"string":
                        if(OSF.OUtil.listContainsValue(param["enum"],arg))
                            break;
                    case"undefined":
                        throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration;
                    default:
                        throw getInvalidParameterString();
                }
            if(param["types"])
                if(!OSF.OUtil.listContainsValue(param["types"],typeof arg))
                    throw getInvalidParameterString();
        }
    };
    this.extractRequiredArguments = function OSF_DDA_ApiMethodCall$ExtractRequiredArguments(userArgs, caller, stateInfo)
    {
        if(userArgs.length < requiredCount)
            throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_MissingRequiredArguments);
        var requiredArgs = [];
        var index;
        for(index = 0; index < requiredCount; index++)
            requiredArgs.push(userArgs[index]);
        this.verifyArguments(requiredParameters,requiredArgs);
        var ret = {};
        for(index = 0; index < requiredCount; index++)
        {
            var param = requiredParameters[index];
            var arg = requiredArgs[index];
            if(param.verify)
            {
                var isValid = param.verify(arg,caller,stateInfo);
                if(!isValid)
                    throw getInvalidParameterString();
            }
            ret[param.name] = arg
        }
        return ret
    },this.fillOptions = function OSF_DDA_ApiMethodCall$FillOptions(options, requiredArgs, caller, stateInfo)
    {
        options = options || {};
        for(var optionName in supportedOptions)
            if(!OSF.OUtil.listContainsKey(options,optionName))
            {
                var value = undefined;
                var option = supportedOptions[optionName];
                if(option.calculate && requiredArgs)
                    value = option.calculate(requiredArgs,caller,stateInfo);
                if(!value && option.defaultValue !== undefined)
                    value = option.defaultValue;
                options[optionName] = value
            }
        return options
    };
    this.constructCallArgs = function OSF_DAA_ApiMethodCall$ConstructCallArgs(required, options, caller, stateInfo)
    {
        var callArgs = {};
        for(var r in required)
            callArgs[r] = required[r];
        for(var o in options)
            callArgs[o] = options[o];
        for(var s in privateStateCallbacks)
            callArgs[s] = privateStateCallbacks[s](caller,stateInfo);
        if(checkCallArgs)
            callArgs = checkCallArgs(callArgs,caller,stateInfo);
        return callArgs
    }
};
OSF.OUtil.setNamespace("AsyncResultEnum",OSF.DDA);
OSF.DDA.AsyncResultEnum.Properties = {
    Context: "Context",
    Value: "Value",
    Status: "Status",
    Error: "Error"
};
Microsoft.Office.WebExtension.AsyncResultStatus = {
    Succeeded: "succeeded",
    Failed: "failed"
};
OSF.DDA.AsyncResultEnum.ErrorCode = {
    Success: 0,
    Failed: 1
};
OSF.DDA.AsyncResultEnum.ErrorProperties = {
    Name: "Name",
    Message: "Message",
    Code: "Code"
};
OSF.DDA.AsyncMethodNames = {};
OSF.DDA.AsyncMethodNames.addNames = function(methodNames)
{
    for(var entry in methodNames)
    {
        var am = {};
        OSF.OUtil.defineEnumerableProperties(am,{
            id: {value: entry},
            displayName: {value: methodNames[entry]}
        });
        OSF.DDA.AsyncMethodNames[entry] = am
    }
};
OSF.DDA.AsyncMethodCall = function OSF_DDA_AsyncMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, onSucceeded, onFailed, checkCallArgs, displayName)
{
    var requiredCount = requiredParameters.length;
    var apiMethods = new OSF.DDA.ApiMethodCall(requiredParameters,supportedOptions,privateStateCallbacks,checkCallArgs,displayName);
    function OSF_DAA_AsyncMethodCall$ExtractOptions(userArgs, requiredArgs, caller, stateInfo)
    {
        if(userArgs.length > requiredCount + 2)
            throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);
        var options,
            parameterCallback;
        for(var i = userArgs.length - 1; i >= requiredCount; i--)
        {
            var argument = userArgs[i];
            switch(typeof argument)
            {
                case"object":
                    if(options)
                        throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);
                    else
                        options = argument;
                    break;
                case"function":
                    if(parameterCallback)
                        throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalFunction);
                    else
                        parameterCallback = argument;
                    break;
                default:
                    throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument);
                    break
            }
        }
        options = apiMethods.fillOptions(options,requiredArgs,caller,stateInfo);
        if(parameterCallback)
            if(options[Microsoft.Office.WebExtension.Parameters.Callback])
                throw Strings.OfficeOM.L_RedundantCallbackSpecification;
            else
                options[Microsoft.Office.WebExtension.Parameters.Callback] = parameterCallback;
        apiMethods.verifyArguments(supportedOptions,options);
        return options
    }
    this.verifyAndExtractCall = function OSF_DAA_AsyncMethodCall$VerifyAndExtractCall(userArgs, caller, stateInfo)
    {
        var required = apiMethods.extractRequiredArguments(userArgs,caller,stateInfo);
        var options = OSF_DAA_AsyncMethodCall$ExtractOptions(userArgs,required,caller,stateInfo);
        var callArgs = apiMethods.constructCallArgs(required,options,caller,stateInfo);
        return callArgs
    };
    this.processResponse = function OSF_DAA_AsyncMethodCall$ProcessResponse(status, response, caller, callArgs)
    {
        var payload;
        if(status == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
            if(onSucceeded)
                payload = onSucceeded(response,caller,callArgs);
            else
                payload = response;
        else if(onFailed)
            payload = onFailed(status,response);
        else
            payload = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
        return payload
    };
    this.getCallArgs = function(suppliedArgs)
    {
        var options,
            parameterCallback;
        for(var i = suppliedArgs.length - 1; i >= requiredCount; i--)
        {
            var argument = suppliedArgs[i];
            switch(typeof argument)
            {
                case"object":
                    options = argument;
                    break;
                case"function":
                    parameterCallback = argument;
                    break
            }
        }
        options = options || {};
        if(parameterCallback)
            options[Microsoft.Office.WebExtension.Parameters.Callback] = parameterCallback;
        return options
    }
};
OSF.DDA.AsyncMethodCallFactory = function()
{
    return{manufacture: function(params)
            {
                var supportedOptions = params.supportedOptions ? OSF.OUtil.createObject(params.supportedOptions) : [];
                var privateStateCallbacks = params.privateStateCallbacks ? OSF.OUtil.createObject(params.privateStateCallbacks) : [];
                return new OSF.DDA.AsyncMethodCall(params.requiredArguments || [],supportedOptions,privateStateCallbacks,params.onSucceeded,params.onFailed,params.checkCallArgs,params.method.displayName)
            }}
}();
OSF.DDA.AsyncMethodCalls = {};
OSF.DDA.AsyncMethodCalls.define = function(callDefinition)
{
    OSF.DDA.AsyncMethodCalls[callDefinition.method.id] = OSF.DDA.AsyncMethodCallFactory.manufacture(callDefinition)
};
OSF.DDA.Error = function OSF_DDA_Error(name, message, code)
{
    OSF.OUtil.defineEnumerableProperties(this,{
        name: {value: name},
        message: {value: message},
        code: {value: code}
    })
};
OSF.DDA.AsyncResult = function OSF_DDA_AsyncResult(initArgs, errorArgs)
{
    OSF.OUtil.defineEnumerableProperties(this,{
        value: {value: initArgs[OSF.DDA.AsyncResultEnum.Properties.Value]},
        status: {value: errorArgs ? Microsoft.Office.WebExtension.AsyncResultStatus.Failed : Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded}
    });
    if(initArgs[OSF.DDA.AsyncResultEnum.Properties.Context])
        OSF.OUtil.defineEnumerableProperty(this,"asyncContext",{value: initArgs[OSF.DDA.AsyncResultEnum.Properties.Context]});
    if(errorArgs)
        OSF.OUtil.defineEnumerableProperty(this,"error",{value: new OSF.DDA.Error(errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name],errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message],errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code])})
};
OSF.DDA.issueAsyncResult = function OSF_DDA$IssueAsyncResult(callArgs, status, payload)
{
    var callback = callArgs[Microsoft.Office.WebExtension.Parameters.Callback];
    if(callback)
    {
        var asyncInitArgs = {};
        asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Context] = callArgs[Microsoft.Office.WebExtension.Parameters.AsyncContext];
        var errorArgs;
        if(status == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
            asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Value] = payload;
        else
        {
            errorArgs = {};
            payload = payload || OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
            errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code] = status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
            errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name] = payload.name || payload;
            errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message] = payload.message || payload
        }
        callback(new OSF.DDA.AsyncResult(asyncInitArgs,errorArgs))
    }
};
OSF.DDA.SyncMethodNames = {};
OSF.DDA.SyncMethodNames.addNames = function(methodNames)
{
    for(var entry in methodNames)
    {
        var am = {};
        OSF.OUtil.defineEnumerableProperties(am,{
            id: {value: entry},
            displayName: {value: methodNames[entry]}
        });
        OSF.DDA.SyncMethodNames[entry] = am
    }
};
OSF.DDA.SyncMethodCall = function OSF_DDA_SyncMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName)
{
    var requiredCount = requiredParameters.length;
    var apiMethods = new OSF.DDA.ApiMethodCall(requiredParameters,supportedOptions,privateStateCallbacks,checkCallArgs,displayName);
    function OSF_DAA_SyncMethodCall$ExtractOptions(userArgs, requiredArgs, caller, stateInfo)
    {
        if(userArgs.length > requiredCount + 1)
            throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);
        var options,
            parameterCallback;
        for(var i = userArgs.length - 1; i >= requiredCount; i--)
        {
            var argument = userArgs[i];
            switch(typeof argument)
            {
                case"object":
                    if(options)
                        throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);
                    else
                        options = argument;
                    break;
                default:
                    throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument);
                    break
            }
        }
        options = apiMethods.fillOptions(options,requiredArgs,caller,stateInfo);
        apiMethods.verifyArguments(supportedOptions,options);
        return options
    }
    this.verifyAndExtractCall = function OSF_DAA_AsyncMethodCall$VerifyAndExtractCall(userArgs, caller, stateInfo)
    {
        var required = apiMethods.extractRequiredArguments(userArgs,caller,stateInfo);
        var options = OSF_DAA_SyncMethodCall$ExtractOptions(userArgs,required,caller,stateInfo);
        var callArgs = apiMethods.constructCallArgs(required,options,caller,stateInfo);
        return callArgs
    }
};
OSF.DDA.SyncMethodCallFactory = function()
{
    return{manufacture: function(params)
            {
                var supportedOptions = params.supportedOptions ? OSF.OUtil.createObject(params.supportedOptions) : [];
                return new OSF.DDA.SyncMethodCall(params.requiredArguments || [],supportedOptions,params.privateStateCallbacks,params.checkCallArgs,params.method.displayName)
            }}
}();
OSF.DDA.SyncMethodCalls = {};
OSF.DDA.SyncMethodCalls.define = function(callDefinition)
{
    OSF.DDA.SyncMethodCalls[callDefinition.method.id] = OSF.DDA.SyncMethodCallFactory.manufacture(callDefinition)
};
OSF.DDA.ListType = function()
{
    var listTypes = {};
    return{
            setListType: function OSF_DDA_ListType$AddListType(t, prop)
            {
                listTypes[t] = prop
            },
            isListType: function OSF_DDA_ListType$IsListType(t)
            {
                return OSF.OUtil.listContainsKey(listTypes,t)
            },
            getDescriptor: function OSF_DDA_ListType$getDescriptor(t)
            {
                return listTypes[t]
            }
        }
}();
OSF.DDA.HostParameterMap = function(specialProcessor, mappings)
{
    var toHostMap = "toHost";
    var fromHostMap = "fromHost";
    var sourceData = "sourceData";
    var self = "self";
    var dynamicTypes = {};
    dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data] = {
        toHost: function(data)
        {
            if(data != null && data.rows !== undefined)
            {
                var tableData = {};
                tableData[OSF.DDA.TableDataProperties.TableRows] = data.rows;
                tableData[OSF.DDA.TableDataProperties.TableHeaders] = data.headers;
                data = tableData
            }
            return data
        },
        fromHost: function(args)
        {
            return args
        }
    };
    dynamicTypes[Microsoft.Office.WebExtension.Parameters.SampleData] = dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data];
    function mapValues(preimageSet, mapping)
    {
        var ret = preimageSet ? {} : undefined;
        for(var entry in preimageSet)
        {
            var preimage = preimageSet[entry];
            var image;
            if(OSF.DDA.ListType.isListType(entry))
            {
                image = [];
                for(var subEntry in preimage)
                    image.push(mapValues(preimage[subEntry],mapping))
            }
            else if(OSF.OUtil.listContainsKey(dynamicTypes,entry))
                image = dynamicTypes[entry][mapping](preimage);
            else if(mapping == fromHostMap && specialProcessor.preserveNesting(entry))
                image = mapValues(preimage,mapping);
            else
            {
                var maps = mappings[entry];
                if(maps)
                {
                    var map = maps[mapping];
                    if(map)
                    {
                        image = map[preimage];
                        if(image === undefined)
                            image = preimage
                    }
                }
                else
                    image = preimage
            }
            ret[entry] = image
        }
        return ret
    }
    function generateArguments(imageSet, parameters)
    {
        var ret;
        for(var param in parameters)
        {
            var arg;
            if(specialProcessor.isComplexType(param))
                arg = generateArguments(imageSet,mappings[param][toHostMap]);
            else
                arg = imageSet[param];
            if(arg != undefined)
            {
                if(!ret)
                    ret = {};
                var index = parameters[param];
                if(index == self)
                    index = param;
                ret[index] = specialProcessor.pack(param,arg)
            }
        }
        return ret
    }
    function extractArguments(source, parameters, extracted)
    {
        if(!extracted)
            extracted = {};
        for(var param in parameters)
        {
            var index = parameters[param];
            var value;
            if(index == self)
                value = source;
            else if(index == sourceData)
            {
                extracted[param] = source.toArray();
                continue
            }
            else
                value = source[index];
            if(value === null || value === undefined)
                extracted[param] = undefined;
            else
            {
                value = specialProcessor.unpack(param,value);
                var map;
                if(specialProcessor.isComplexType(param))
                {
                    map = mappings[param][fromHostMap];
                    if(specialProcessor.preserveNesting(param))
                        extracted[param] = extractArguments(value,map);
                    else
                        extractArguments(value,map,extracted)
                }
                else if(OSF.DDA.ListType.isListType(param))
                {
                    map = {};
                    var entryDescriptor = OSF.DDA.ListType.getDescriptor(param);
                    map[entryDescriptor] = self;
                    var extractedValues = new Array(value.length);
                    for(var item in value)
                        extractedValues[item] = extractArguments(value[item],map);
                    extracted[param] = extractedValues
                }
                else
                    extracted[param] = value
            }
        }
        return extracted
    }
    function applyMap(mapName, preimage, mapping)
    {
        var parameters = mappings[mapName][mapping];
        var image;
        if(mapping == "toHost")
        {
            var imageSet = mapValues(preimage,mapping);
            image = generateArguments(imageSet,parameters)
        }
        else if(mapping == "fromHost")
        {
            var argumentSet = extractArguments(preimage,parameters);
            image = mapValues(argumentSet,mapping)
        }
        return image
    }
    if(!mappings)
        mappings = {};
    this.addMapping = function(mapName, description)
    {
        var toHost,
            fromHost;
        if(description.map)
        {
            toHost = description.map;
            fromHost = {};
            for(var preimage in toHost)
            {
                var image = toHost[preimage];
                if(image == self)
                    image = preimage;
                fromHost[image] = preimage
            }
        }
        else
        {
            toHost = description.toHost;
            fromHost = description.fromHost
        }
        var pair = mappings[mapName];
        if(pair)
        {
            var currMap = pair[toHostMap];
            for(var th in currMap)
                toHost[th] = currMap[th];
            currMap = pair[fromHostMap];
            for(var fh in currMap)
                fromHost[fh] = currMap[fh]
        }
        else
            pair = mappings[mapName] = {};
        pair[toHostMap] = toHost;
        pair[fromHostMap] = fromHost
    };
    this.toHost = function(mapName, preimage)
    {
        return applyMap(mapName,preimage,toHostMap)
    };
    this.fromHost = function(mapName, image)
    {
        return applyMap(mapName,image,fromHostMap)
    };
    this.self = self;
    this.sourceData = sourceData;
    this.addComplexType = function(ct)
    {
        specialProcessor.addComplexType(ct)
    };
    this.getDynamicType = function(dt)
    {
        return specialProcessor.getDynamicType(dt)
    };
    this.setDynamicType = function(dt, handler)
    {
        specialProcessor.setDynamicType(dt,handler)
    };
    this.dynamicTypes = dynamicTypes;
    this.doMapValues = function(preimageSet, mapping)
    {
        return mapValues(preimageSet,mapping)
    }
};
OSF.DDA.SpecialProcessor = function(complexTypes, dynamicTypes)
{
    this.addComplexType = function OSF_DDA_SpecialProcessor$addComplexType(ct)
    {
        complexTypes.push(ct)
    };
    this.getDynamicType = function OSF_DDA_SpecialProcessor$getDynamicType(dt)
    {
        return dynamicTypes[dt]
    };
    this.setDynamicType = function OSF_DDA_SpecialProcessor$setDynamicType(dt, handler)
    {
        dynamicTypes[dt] = handler
    };
    this.isComplexType = function OSF_DDA_SpecialProcessor$isComplexType(t)
    {
        return OSF.OUtil.listContainsValue(complexTypes,t)
    };
    this.isDynamicType = function OSF_DDA_SpecialProcessor$isDynamicType(p)
    {
        return OSF.OUtil.listContainsKey(dynamicTypes,p)
    };
    this.preserveNesting = function OSF_DDA_SpecialProcessor$preserveNesting(p)
    {
        var pn = [];
        if(OSF.DDA.PropertyDescriptors)
            pn.push(OSF.DDA.PropertyDescriptors.Subset);
        if(OSF.DDA.DataNodeEventProperties)
            pn = pn.concat([OSF.DDA.DataNodeEventProperties.OldNode,OSF.DDA.DataNodeEventProperties.NewNode,OSF.DDA.DataNodeEventProperties.NextSiblingNode]);
        return OSF.OUtil.listContainsValue(pn,p)
    };
    this.pack = function OSF_DDA_SpecialProcessor$pack(param, arg)
    {
        var value;
        if(this.isDynamicType(param))
            value = dynamicTypes[param].toHost(arg);
        else
            value = arg;
        return value
    };
    this.unpack = function OSF_DDA_SpecialProcessor$unpack(param, arg)
    {
        var value;
        if(this.isDynamicType(param))
            value = dynamicTypes[param].fromHost(arg);
        else
            value = arg;
        return value
    }
};
OSF.DDA.getDecoratedParameterMap = function(specialProcessor, initialDefs)
{
    var parameterMap = new OSF.DDA.HostParameterMap(specialProcessor);
    var self = parameterMap.self;
    function createObject(properties)
    {
        var obj = null;
        if(properties)
        {
            obj = {};
            var len = properties.length;
            for(var i = 0; i < len; i++)
                obj[properties[i].name] = properties[i].value
        }
        return obj
    }
    parameterMap.define = function define(definition)
    {
        var args = {};
        var toHost = createObject(definition.toHost);
        if(definition.invertible)
            args.map = toHost;
        else if(definition.canonical)
            args.toHost = args.fromHost = toHost;
        else
        {
            args.toHost = toHost;
            args.fromHost = createObject(definition.fromHost)
        }
        parameterMap.addMapping(definition.type,args);
        if(definition.isComplexType)
            parameterMap.addComplexType(definition.type)
    };
    for(var id in initialDefs)
        parameterMap.define(initialDefs[id]);
    return parameterMap
};
OSF.OUtil.setNamespace("DispIdHost",OSF.DDA);
OSF.DDA.DispIdHost.Methods = {
    InvokeMethod: "invokeMethod",
    AddEventHandler: "addEventHandler",
    RemoveEventHandler: "removeEventHandler",
    OpenDialog: "openDialog",
    CloseDialog: "closeDialog",
    MessageParent: "messageParent",
    SendMessage: "sendMessage"
};
OSF.DDA.DispIdHost.Delegates = {
    ExecuteAsync: "executeAsync",
    RegisterEventAsync: "registerEventAsync",
    UnregisterEventAsync: "unregisterEventAsync",
    ParameterMap: "parameterMap",
    OpenDialog: "openDialog",
    CloseDialog: "closeDialog",
    MessageParent: "messageParent",
    SendMessage: "sendMessage"
};
OSF.DDA.DispIdHost.Facade = function OSF_DDA_DispIdHost_Facade(getDelegateMethods, parameterMap)
{
    var dispIdMap = {};
    var jsom = OSF.DDA.AsyncMethodNames;
    var did = OSF.DDA.MethodDispId;
    var methodMap = {
            GoToByIdAsync: did.dispidNavigateToMethod,
            GetSelectedDataAsync: did.dispidGetSelectedDataMethod,
            SetSelectedDataAsync: did.dispidSetSelectedDataMethod,
            GetDocumentCopyChunkAsync: did.dispidGetDocumentCopyChunkMethod,
            ReleaseDocumentCopyAsync: did.dispidReleaseDocumentCopyMethod,
            GetDocumentCopyAsync: did.dispidGetDocumentCopyMethod,
            AddFromSelectionAsync: did.dispidAddBindingFromSelectionMethod,
            AddFromPromptAsync: did.dispidAddBindingFromPromptMethod,
            AddFromNamedItemAsync: did.dispidAddBindingFromNamedItemMethod,
            GetAllAsync: did.dispidGetAllBindingsMethod,
            GetByIdAsync: did.dispidGetBindingMethod,
            ReleaseByIdAsync: did.dispidReleaseBindingMethod,
            GetDataAsync: did.dispidGetBindingDataMethod,
            SetDataAsync: did.dispidSetBindingDataMethod,
            AddRowsAsync: did.dispidAddRowsMethod,
            AddColumnsAsync: did.dispidAddColumnsMethod,
            DeleteAllDataValuesAsync: did.dispidClearAllRowsMethod,
            RefreshAsync: did.dispidLoadSettingsMethod,
            SaveAsync: did.dispidSaveSettingsMethod,
            GetActiveViewAsync: did.dispidGetActiveViewMethod,
            GetFilePropertiesAsync: did.dispidGetFilePropertiesMethod,
            GetOfficeThemeAsync: did.dispidGetOfficeThemeMethod,
            GetDocumentThemeAsync: did.dispidGetDocumentThemeMethod,
            ClearFormatsAsync: did.dispidClearFormatsMethod,
            SetTableOptionsAsync: did.dispidSetTableOptionsMethod,
            SetFormatsAsync: did.dispidSetFormatsMethod,
            GetAccessTokenAsync: did.dispidGetAccessTokenMethod,
            GetAuthContextAsync: did.dispidGetAuthContextMethod,
            ExecuteRichApiRequestAsync: did.dispidExecuteRichApiRequestMethod,
            AppCommandInvocationCompletedAsync: did.dispidAppCommandInvocationCompletedMethod,
            CloseContainerAsync: did.dispidCloseContainerMethod,
            OpenBrowserWindow: did.dispidOpenBrowserWindow,
            CreateDocumentAsync: did.dispidCreateDocumentMethod,
            InsertFormAsync: did.dispidInsertFormMethod,
            ExecuteFeature: did.dispidExecuteFeature,
            QueryFeature: did.dispidQueryFeature,
            AddDataPartAsync: did.dispidAddDataPartMethod,
            GetDataPartByIdAsync: did.dispidGetDataPartByIdMethod,
            GetDataPartsByNameSpaceAsync: did.dispidGetDataPartsByNamespaceMethod,
            GetPartXmlAsync: did.dispidGetDataPartXmlMethod,
            GetPartNodesAsync: did.dispidGetDataPartNodesMethod,
            DeleteDataPartAsync: did.dispidDeleteDataPartMethod,
            GetNodeValueAsync: did.dispidGetDataNodeValueMethod,
            GetNodeXmlAsync: did.dispidGetDataNodeXmlMethod,
            GetRelativeNodesAsync: did.dispidGetDataNodesMethod,
            SetNodeValueAsync: did.dispidSetDataNodeValueMethod,
            SetNodeXmlAsync: did.dispidSetDataNodeXmlMethod,
            AddDataPartNamespaceAsync: did.dispidAddDataNamespaceMethod,
            GetDataPartNamespaceAsync: did.dispidGetDataUriByPrefixMethod,
            GetDataPartPrefixAsync: did.dispidGetDataPrefixByUriMethod,
            GetNodeTextAsync: did.dispidGetDataNodeTextMethod,
            SetNodeTextAsync: did.dispidSetDataNodeTextMethod,
            GetSelectedTask: did.dispidGetSelectedTaskMethod,
            GetTask: did.dispidGetTaskMethod,
            GetWSSUrl: did.dispidGetWSSUrlMethod,
            GetTaskField: did.dispidGetTaskFieldMethod,
            GetSelectedResource: did.dispidGetSelectedResourceMethod,
            GetResourceField: did.dispidGetResourceFieldMethod,
            GetProjectField: did.dispidGetProjectFieldMethod,
            GetSelectedView: did.dispidGetSelectedViewMethod,
            GetTaskByIndex: did.dispidGetTaskByIndexMethod,
            GetResourceByIndex: did.dispidGetResourceByIndexMethod,
            SetTaskField: did.dispidSetTaskFieldMethod,
            SetResourceField: did.dispidSetResourceFieldMethod,
            GetMaxTaskIndex: did.dispidGetMaxTaskIndexMethod,
            GetMaxResourceIndex: did.dispidGetMaxResourceIndexMethod,
            CreateTask: did.dispidCreateTaskMethod
        };
    for(var method in methodMap)
        if(jsom[method])
            dispIdMap[jsom[method].id] = methodMap[method];
    jsom = OSF.DDA.SyncMethodNames;
    did = OSF.DDA.MethodDispId;
    var syncMethodMap = {
            MessageParent: did.dispidMessageParentMethod,
            SendMessage: did.dispidSendMessageMethod
        };
    for(var method in syncMethodMap)
        if(jsom[method])
            dispIdMap[jsom[method].id] = syncMethodMap[method];
    jsom = Microsoft.Office.WebExtension.EventType;
    did = OSF.DDA.EventDispId;
    var eventMap = {
            SettingsChanged: did.dispidSettingsChangedEvent,
            DocumentSelectionChanged: did.dispidDocumentSelectionChangedEvent,
            BindingSelectionChanged: did.dispidBindingSelectionChangedEvent,
            BindingDataChanged: did.dispidBindingDataChangedEvent,
            ActiveViewChanged: did.dispidActiveViewChangedEvent,
            OfficeThemeChanged: did.dispidOfficeThemeChangedEvent,
            DocumentThemeChanged: did.dispidDocumentThemeChangedEvent,
            AppCommandInvoked: did.dispidAppCommandInvokedEvent,
            DialogMessageReceived: did.dispidDialogMessageReceivedEvent,
            DialogParentMessageReceived: did.dispidDialogParentMessageReceivedEvent,
            ObjectDeleted: did.dispidObjectDeletedEvent,
            ObjectSelectionChanged: did.dispidObjectSelectionChangedEvent,
            ObjectDataChanged: did.dispidObjectDataChangedEvent,
            ContentControlAdded: did.dispidContentControlAddedEvent,
            RichApiMessage: did.dispidRichApiMessageEvent,
            ItemChanged: did.dispidOlkItemSelectedChangedEvent,
            RecipientsChanged: did.dispidOlkRecipientsChangedEvent,
            AppointmentTimeChanged: did.dispidOlkAppointmentTimeChangedEvent,
            RecurrenceChanged: did.dispidOlkRecurrenceChangedEvent,
            AttachmentsChanged: did.dispidOlkAttachmentsChangedEvent,
            EnhancedLocationsChanged: did.dispidOlkEnhancedLocationsChangedEvent,
            InfobarClicked: did.dispidOlkInfobarClickedEvent,
            TaskSelectionChanged: did.dispidTaskSelectionChangedEvent,
            ResourceSelectionChanged: did.dispidResourceSelectionChangedEvent,
            ViewSelectionChanged: did.dispidViewSelectionChangedEvent,
            DataNodeInserted: did.dispidDataNodeAddedEvent,
            DataNodeReplaced: did.dispidDataNodeReplacedEvent,
            DataNodeDeleted: did.dispidDataNodeDeletedEvent
        };
    for(var event in eventMap)
        if(jsom[event])
            dispIdMap[jsom[event]] = eventMap[event];
    function IsObjectEvent(dispId)
    {
        return dispId == OSF.DDA.EventDispId.dispidObjectDeletedEvent || dispId == OSF.DDA.EventDispId.dispidObjectSelectionChangedEvent || dispId == OSF.DDA.EventDispId.dispidObjectDataChangedEvent || dispId == OSF.DDA.EventDispId.dispidContentControlAddedEvent
    }
    function onException(ex, asyncMethodCall, suppliedArgs, callArgs)
    {
        if(typeof ex == "number")
        {
            if(!callArgs)
                callArgs = asyncMethodCall.getCallArgs(suppliedArgs);
            OSF.DDA.issueAsyncResult(callArgs,ex,OSF.DDA.ErrorCodeManager.getErrorArgs(ex))
        }
        else
            throw ex;
    }
    this[OSF.DDA.DispIdHost.Methods.InvokeMethod] = function OSF_DDA_DispIdHost_Facade$InvokeMethod(method, suppliedArguments, caller, privateState)
    {
        var callArgs;
        try
        {
            var methodName = method.id;
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[methodName];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments,caller,privateState);
            var dispId = dispIdMap[methodName];
            var delegate = getDelegateMethods(methodName);
            var richApiInExcelMethodSubstitution = null;
            if(window.Excel && window.Office.context.requirements.isSetSupported("RedirectV1Api"))
                window.Excel._RedirectV1APIs = true;
            if(window.Excel && window.Excel._RedirectV1APIs && (richApiInExcelMethodSubstitution = window.Excel._V1APIMap[methodName]))
            {
                var preprocessedCallArgs = OSF.OUtil.shallowCopy(callArgs);
                delete preprocessedCallArgs[Microsoft.Office.WebExtension.Parameters.AsyncContext];
                if(richApiInExcelMethodSubstitution.preprocess)
                    preprocessedCallArgs = richApiInExcelMethodSubstitution.preprocess(preprocessedCallArgs);
                var ctx = new window.Excel.RequestContext;
                var result = richApiInExcelMethodSubstitution.call(ctx,preprocessedCallArgs);
                ctx.sync().then(function()
                {
                    var response = result.value;
                    var status = response.status;
                    delete response["status"];
                    delete response["@odata.type"];
                    if(richApiInExcelMethodSubstitution.postprocess)
                        response = richApiInExcelMethodSubstitution.postprocess(response,preprocessedCallArgs);
                    if(status != 0)
                        response = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
                    OSF.DDA.issueAsyncResult(callArgs,status,response)
                })["catch"](function(error)
                {
                    OSF.DDA.issueAsyncResult(callArgs,OSF.DDA.ErrorCodeManager.errorCodes.ooeFailure,null)
                })
            }
            else
            {
                var hostCallArgs;
                if(parameterMap.toHost)
                    hostCallArgs = parameterMap.toHost(dispId,callArgs);
                else
                    hostCallArgs = callArgs;
                var startTime = (new Date).getTime();
                delegate[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]({
                    dispId: dispId,
                    hostCallArgs: hostCallArgs,
                    onCalling: function OSF_DDA_DispIdFacade$Execute_onCalling(){},
                    onReceiving: function OSF_DDA_DispIdFacade$Execute_onReceiving(){},
                    onComplete: function(status, hostResponseArgs)
                    {
                        var responseArgs;
                        if(status == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
                            if(parameterMap.fromHost)
                                responseArgs = parameterMap.fromHost(dispId,hostResponseArgs);
                            else
                                responseArgs = hostResponseArgs;
                        else
                            responseArgs = hostResponseArgs;
                        var payload = asyncMethodCall.processResponse(status,responseArgs,caller,callArgs);
                        OSF.DDA.issueAsyncResult(callArgs,status,payload);
                        if(OSF.AppTelemetry && !(OSF.ConstantNames && OSF.ConstantNames.IsCustomFunctionsRuntime))
                            OSF.AppTelemetry.onMethodDone(dispId,hostCallArgs,Math.abs((new Date).getTime() - startTime),status)
                    }
                })
            }
        }
        catch(ex)
        {
            onException(ex,asyncMethodCall,suppliedArguments,callArgs)
        }
    };
    this[OSF.DDA.DispIdHost.Methods.AddEventHandler] = function OSF_DDA_DispIdHost_Facade$AddEventHandler(suppliedArguments, eventDispatch, caller, isPopupWindow)
    {
        var callArgs;
        var eventType,
            handler;
        var isObjectEvent = false;
        function onEnsureRegistration(status)
        {
            if(status == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
            {
                var added = !isObjectEvent ? eventDispatch.addEventHandler(eventType,handler) : eventDispatch.addObjectEventHandler(eventType,callArgs[Microsoft.Office.WebExtension.Parameters.Id],handler);
                if(!added)
                    status = OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerAdditionFailed
            }
            var error;
            if(status != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
                error = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
            OSF.DDA.issueAsyncResult(callArgs,status,error)
        }
        try
        {
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.AddHandlerAsync.id];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments,caller,eventDispatch);
            eventType = callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
            handler = callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
            if(isPopupWindow)
            {
                onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
                return
            }
            var dispId = dispIdMap[eventType];
            isObjectEvent = IsObjectEvent(dispId);
            var targetId = isObjectEvent ? callArgs[Microsoft.Office.WebExtension.Parameters.Id] : caller.id || "";
            var count = isObjectEvent ? eventDispatch.getObjectEventHandlerCount(eventType,targetId) : eventDispatch.getEventHandlerCount(eventType);
            if(count == 0)
            {
                var invoker = getDelegateMethods(eventType)[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];
                invoker({
                    eventType: eventType,
                    dispId: dispId,
                    targetId: targetId,
                    onCalling: function OSF_DDA_DispIdFacade$Execute_onCalling()
                    {
                        OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)
                    },
                    onReceiving: function OSF_DDA_DispIdFacade$Execute_onReceiving()
                    {
                        OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)
                    },
                    onComplete: onEnsureRegistration,
                    onEvent: function handleEvent(hostArgs)
                    {
                        var args = parameterMap.fromHost(dispId,hostArgs);
                        if(!isObjectEvent)
                            eventDispatch.fireEvent(OSF.DDA.OMFactory.manufactureEventArgs(eventType,caller,args));
                        else
                            eventDispatch.fireObjectEvent(targetId,OSF.DDA.OMFactory.manufactureEventArgs(eventType,targetId,args))
                    }
                })
            }
            else
                onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
        }
        catch(ex)
        {
            onException(ex,asyncMethodCall,suppliedArguments,callArgs)
        }
    };
    this[OSF.DDA.DispIdHost.Methods.RemoveEventHandler] = function OSF_DDA_DispIdHost_Facade$RemoveEventHandler(suppliedArguments, eventDispatch, caller)
    {
        var callArgs;
        var eventType,
            handler;
        var isObjectEvent = false;
        function onEnsureRegistration(status)
        {
            var error;
            if(status != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
                error = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
            OSF.DDA.issueAsyncResult(callArgs,status,error)
        }
        try
        {
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.id];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments,caller,eventDispatch);
            eventType = callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
            handler = callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
            var dispId = dispIdMap[eventType];
            isObjectEvent = IsObjectEvent(dispId);
            var targetId = isObjectEvent ? callArgs[Microsoft.Office.WebExtension.Parameters.Id] : caller.id || "";
            var status,
                removeSuccess;
            if(handler === null)
            {
                removeSuccess = isObjectEvent ? eventDispatch.clearObjectEventHandlers(eventType,targetId) : eventDispatch.clearEventHandlers(eventType);
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess
            }
            else
            {
                removeSuccess = isObjectEvent ? eventDispatch.removeObjectEventHandler(eventType,targetId,handler) : eventDispatch.removeEventHandler(eventType,handler);
                status = removeSuccess ? OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess : OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist
            }
            var count = isObjectEvent ? eventDispatch.getObjectEventHandlerCount(eventType,targetId) : eventDispatch.getEventHandlerCount(eventType);
            if(removeSuccess && count == 0)
            {
                var invoker = getDelegateMethods(eventType)[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];
                invoker({
                    eventType: eventType,
                    dispId: dispId,
                    targetId: targetId,
                    onCalling: function OSF_DDA_DispIdFacade$Execute_onCalling()
                    {
                        OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)
                    },
                    onReceiving: function OSF_DDA_DispIdFacade$Execute_onReceiving()
                    {
                        OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)
                    },
                    onComplete: onEnsureRegistration
                })
            }
            else
                onEnsureRegistration(status)
        }
        catch(ex)
        {
            onException(ex,asyncMethodCall,suppliedArguments,callArgs)
        }
    };
    this[OSF.DDA.DispIdHost.Methods.OpenDialog] = function OSF_DDA_DispIdHost_Facade$OpenDialog(suppliedArguments, eventDispatch, caller)
    {
        var callArgs;
        var targetId;
        var dialogMessageEvent = Microsoft.Office.WebExtension.EventType.DialogMessageReceived;
        var dialogOtherEvent = Microsoft.Office.WebExtension.EventType.DialogEventReceived;
        function onEnsureRegistration(status)
        {
            var payload;
            if(status != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
                payload = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
            else
            {
                var onSucceedArgs = {};
                onSucceedArgs[Microsoft.Office.WebExtension.Parameters.Id] = targetId;
                onSucceedArgs[Microsoft.Office.WebExtension.Parameters.Data] = eventDispatch;
                var payload = asyncMethodCall.processResponse(status,onSucceedArgs,caller,callArgs);
                OSF.DialogShownStatus.hasDialogShown = true;
                eventDispatch.clearEventHandlers(dialogMessageEvent);
                eventDispatch.clearEventHandlers(dialogOtherEvent)
            }
            OSF.DDA.issueAsyncResult(callArgs,status,payload)
        }
        try
        {
            if(dialogMessageEvent == undefined || dialogOtherEvent == undefined)
                onEnsureRegistration(OSF.DDA.ErrorCodeManager.ooeOperationNotSupported);
            if(OSF.DDA.AsyncMethodNames.DisplayDialogAsync == null)
            {
                onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
                return
            }
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.DisplayDialogAsync.id];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments,caller,eventDispatch);
            var dispId = dispIdMap[dialogMessageEvent];
            var delegateMethods = getDelegateMethods(dialogMessageEvent);
            var invoker = delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog] != undefined ? delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog] : delegateMethods[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];
            targetId = JSON.stringify(callArgs);
            if(!OSF.DialogShownStatus.hasDialogShown)
            {
                eventDispatch.clearQueuedEvent(dialogMessageEvent);
                eventDispatch.clearQueuedEvent(dialogOtherEvent);
                eventDispatch.clearQueuedEvent(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived)
            }
            invoker({
                eventType: dialogMessageEvent,
                dispId: dispId,
                targetId: targetId,
                onCalling: function OSF_DDA_DispIdFacade$Execute_onCalling()
                {
                    OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)
                },
                onReceiving: function OSF_DDA_DispIdFacade$Execute_onReceiving()
                {
                    OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)
                },
                onComplete: onEnsureRegistration,
                onEvent: function handleEvent(hostArgs)
                {
                    var args = parameterMap.fromHost(dispId,hostArgs);
                    var event = OSF.DDA.OMFactory.manufactureEventArgs(dialogMessageEvent,caller,args);
                    if(event.type == dialogOtherEvent)
                    {
                        var payload = OSF.DDA.ErrorCodeManager.getErrorArgs(event.error);
                        var errorArgs = {};
                        errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code] = status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                        errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name] = payload.name || payload;
                        errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message] = payload.message || payload;
                        event.error = new OSF.DDA.Error(errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name],errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message],errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code])
                    }
                    eventDispatch.fireOrQueueEvent(event);
                    if(args[OSF.DDA.PropertyDescriptors.MessageType] == OSF.DialogMessageType.DialogClosed)
                    {
                        eventDispatch.clearEventHandlers(dialogMessageEvent);
                        eventDispatch.clearEventHandlers(dialogOtherEvent);
                        eventDispatch.clearEventHandlers(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived);
                        OSF.DialogShownStatus.hasDialogShown = false
                    }
                }
            })
        }
        catch(ex)
        {
            onException(ex,asyncMethodCall,suppliedArguments,callArgs)
        }
    };
    this[OSF.DDA.DispIdHost.Methods.CloseDialog] = function OSF_DDA_DispIdHost_Facade$CloseDialog(suppliedArguments, targetId, eventDispatch, caller)
    {
        var callArgs;
        var dialogMessageEvent,
            dialogOtherEvent;
        var closeStatus = OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
        function closeCallback(status)
        {
            closeStatus = status;
            OSF.DialogShownStatus.hasDialogShown = false
        }
        try
        {
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.CloseAsync.id];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments,caller,eventDispatch);
            dialogMessageEvent = Microsoft.Office.WebExtension.EventType.DialogMessageReceived;
            dialogOtherEvent = Microsoft.Office.WebExtension.EventType.DialogEventReceived;
            eventDispatch.clearEventHandlers(dialogMessageEvent);
            eventDispatch.clearEventHandlers(dialogOtherEvent);
            var dispId = dispIdMap[dialogMessageEvent];
            var delegateMethods = getDelegateMethods(dialogMessageEvent);
            var invoker = delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog] != undefined ? delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog] : delegateMethods[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];
            invoker({
                eventType: dialogMessageEvent,
                dispId: dispId,
                targetId: targetId,
                onCalling: function OSF_DDA_DispIdFacade$Execute_onCalling()
                {
                    OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)
                },
                onReceiving: function OSF_DDA_DispIdFacade$Execute_onReceiving()
                {
                    OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)
                },
                onComplete: closeCallback
            })
        }
        catch(ex)
        {
            onException(ex,asyncMethodCall,suppliedArguments,callArgs)
        }
        if(closeStatus != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
            throw OSF.OUtil.formatString(Strings.OfficeOM.L_FunctionCallFailed,OSF.DDA.AsyncMethodNames.CloseAsync.displayName,closeStatus);
    };
    this[OSF.DDA.DispIdHost.Methods.MessageParent] = function OSF_DDA_DispIdHost_Facade$MessageParent(suppliedArguments, caller)
    {
        var stateInfo = {};
        var syncMethodCall = OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.MessageParent.id];
        var callArgs = syncMethodCall.verifyAndExtractCall(suppliedArguments,caller,stateInfo);
        var delegate = getDelegateMethods(OSF.DDA.SyncMethodNames.MessageParent.id);
        var invoker = delegate[OSF.DDA.DispIdHost.Delegates.MessageParent];
        var dispId = dispIdMap[OSF.DDA.SyncMethodNames.MessageParent.id];
        return invoker({
                dispId: dispId,
                hostCallArgs: callArgs,
                onCalling: function OSF_DDA_DispIdFacade$Execute_onCalling()
                {
                    OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)
                },
                onReceiving: function OSF_DDA_DispIdFacade$Execute_onReceiving()
                {
                    OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)
                }
            })
    };
    this[OSF.DDA.DispIdHost.Methods.SendMessage] = function OSF_DDA_DispIdHost_Facade$SendMessage(suppliedArguments, eventDispatch, caller)
    {
        var stateInfo = {};
        var syncMethodCall = OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.SendMessage.id];
        var callArgs = syncMethodCall.verifyAndExtractCall(suppliedArguments,caller,stateInfo);
        var delegate = getDelegateMethods(OSF.DDA.SyncMethodNames.SendMessage.id);
        var invoker = delegate[OSF.DDA.DispIdHost.Delegates.SendMessage];
        var dispId = dispIdMap[OSF.DDA.SyncMethodNames.SendMessage.id];
        return invoker({
                dispId: dispId,
                hostCallArgs: callArgs,
                onCalling: function OSF_DDA_DispIdFacade$Execute_onCalling()
                {
                    OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)
                },
                onReceiving: function OSF_DDA_DispIdFacade$Execute_onReceiving()
                {
                    OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)
                }
            })
    }
};
OSF.DDA.DispIdHost.addAsyncMethods = function OSF_DDA_DispIdHost$AddAsyncMethods(target, asyncMethodNames, privateState)
{
    for(var entry in asyncMethodNames)
    {
        var method = asyncMethodNames[entry];
        var name = method.displayName;
        if(!target[name])
            OSF.OUtil.defineEnumerableProperty(target,name,{value: function(asyncMethod)
                {
                    return function()
                        {
                            var invokeMethod = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.InvokeMethod];
                            invokeMethod(asyncMethod,arguments,target,privateState)
                        }
                }(method)})
    }
};
OSF.DDA.DispIdHost.addEventSupport = function OSF_DDA_DispIdHost$AddEventSupport(target, eventDispatch, isPopupWindow)
{
    var add = OSF.DDA.AsyncMethodNames.AddHandlerAsync.displayName;
    var remove = OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.displayName;
    if(!target[add])
        OSF.OUtil.defineEnumerableProperty(target,add,{value: function()
            {
                var addEventHandler = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.AddEventHandler];
                addEventHandler(arguments,eventDispatch,target,isPopupWindow)
            }});
    if(!target[remove])
        OSF.OUtil.defineEnumerableProperty(target,remove,{value: function()
            {
                var removeEventHandler = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.RemoveEventHandler];
                removeEventHandler(arguments,eventDispatch,target)
            }})
};
OSF.ShowWindowDialogParameterKeys = {
    Url: "url",
    Width: "width",
    Height: "height",
    DisplayInIframe: "displayInIframe",
    HideTitle: "hideTitle",
    UseDeviceIndependentPixels: "useDeviceIndependentPixels",
    PromptBeforeOpen: "promptBeforeOpen",
    EnforceAppDomain: "enforceAppDomain"
};
OSF.HostThemeButtonStyleKeys = {
    ButtonBorderColor: "buttonBorderColor",
    ButtonBackgroundColor: "buttonBackgroundColor"
};
OSF.OmexPageParameterKeys = {
    AppName: "client",
    AppVersion: "cv",
    AppUILocale: "ui",
    AppDomain: "appDomain",
    StoreLocator: "rs",
    AssetId: "assetid",
    NotificationType: "notificationType",
    AppCorrelationId: "corr",
    AuthType: "authType",
    AppId: "appid",
    Scopes: "scopes"
};
OSF.AuthType = {
    Anonymous: 0,
    MSA: 1,
    OrgId: 2,
    ADAL: 3
};
OSF.OmexMessageKeys = {
    MessageType: "messageType",
    MessageValue: "messageValue"
};
OSF.OmexRemoveAddinMessageKeys = {
    RemoveAddinResultCode: "resultCode",
    RemoveAddinResultValue: "resultValue"
};
OSF.OmexRemoveAddinResultCode = {
    Success: 0,
    ClientError: 400,
    ServerError: 500,
    UnknownError: 600
};
var OfficeExt;
(function(OfficeExt)
{
    var WACUtils;
    (function(WACUtils)
    {
        var _trustedDomain = "^https://[a-z0-9-]+.(officeapps.live|officeapps-df.live|partner.officewebapps).com/";
        function parseAppContextFromWindowName(skipSessionStorage, windowName)
        {
            return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage,windowName,OSF.WindowNameItemKeys.AppContext)
        }
        WACUtils.parseAppContextFromWindowName = parseAppContextFromWindowName;
        function serializeObjectToString(obj)
        {
            if(typeof JSON !== "undefined")
                try
                {
                    return JSON.stringify(obj)
                }
                catch(ex){}
            return""
        }
        WACUtils.serializeObjectToString = serializeObjectToString;
        function isHostTrusted()
        {
            return new RegExp(_trustedDomain).test(OSF.getClientEndPoint()._targetUrl.toLowerCase())
        }
        WACUtils.isHostTrusted = isHostTrusted;
        function addHostInfoAsQueryParam(url, hostInfoValue)
        {
            if(!url)
                return null;
            url = url.trim() || "";
            var questionMark = "?";
            var hostInfo = "_host_Info=";
            var ampHostInfo = "&_host_Info=";
            var fragmentSeparator = "#";
            var urlParts = url.split(fragmentSeparator);
            var urlWithoutFragment = urlParts.shift();
            var fragment = urlParts.join(fragmentSeparator);
            var querySplits = urlWithoutFragment.split(questionMark);
            var urlWithoutFragmentWithHostInfo;
            if(querySplits.length > 1)
                urlWithoutFragmentWithHostInfo = urlWithoutFragment + ampHostInfo + hostInfoValue;
            else if(querySplits.length > 0)
                urlWithoutFragmentWithHostInfo = urlWithoutFragment + questionMark + hostInfo + hostInfoValue;
            if(fragment)
                return[urlWithoutFragmentWithHostInfo,fragmentSeparator,fragment].join("");
            else
                return urlWithoutFragmentWithHostInfo
        }
        WACUtils.addHostInfoAsQueryParam = addHostInfoAsQueryParam;
        function getDomainForUrl(url)
        {
            if(!url)
                return null;
            var url_parser = document.createElement("a");
            url_parser.href = url;
            return url_parser.protocol + "//" + url_parser.host
        }
        WACUtils.getDomainForUrl = getDomainForUrl;
        function shouldUseLocalStorageToPassMessage()
        {
            try
            {
                var osList = ["Windows NT 6.1","Windows NT 6.2","Windows NT 6.3"];
                var userAgent = window.navigator.userAgent;
                for(var i = 0, len = osList.length; i < len; i++)
                    if(userAgent.indexOf(osList[i]) > -1)
                        return isInternetExplorer();
                return false
            }
            catch(e)
            {
                logExceptionToBrowserConsole("Error happens in shouldUseLocalStorageToPassMessage.",e);
                return false
            }
        }
        WACUtils.shouldUseLocalStorageToPassMessage = shouldUseLocalStorageToPassMessage;
        function isInternetExplorer()
        {
            try
            {
                var userAgent = window.navigator.userAgent;
                return userAgent.indexOf("MSIE ") > -1 || userAgent.indexOf("Trident/") > -1 || userAgent.indexOf("Edge/") > -1
            }
            catch(e)
            {
                logExceptionToBrowserConsole("Error happens in isInternetExplorer.",e);
                return false
            }
        }
        WACUtils.isInternetExplorer = isInternetExplorer;
        function removeMatchesFromLocalStorage(regexPatterns)
        {
            var _localStorage = OSF.OUtil.getLocalStorage();
            var keys = _localStorage.getKeysWithPrefix("");
            for(var i = 0, len = keys.length; i < len; i++)
            {
                var key = keys[i];
                for(var j = 0, lenRegex = regexPatterns.length; j < lenRegex; j++)
                    if(regexPatterns[j].test(key))
                    {
                        _localStorage.removeItem(key);
                        break
                    }
            }
        }
        WACUtils.removeMatchesFromLocalStorage = removeMatchesFromLocalStorage;
        function logExceptionToBrowserConsole(message, exception)
        {
            OsfMsAjaxFactory.msAjaxDebug.trace(message + " Exception details: " + serializeObjectToString(exception))
        }
        WACUtils.logExceptionToBrowserConsole = logExceptionToBrowserConsole;
        var CacheConstants = function()
            {
                function CacheConstants(){}
                CacheConstants.GatedCacheKeyPrefix = "__OSF_GATED_OMEX.";
                CacheConstants.AuthenticatedAppInstallInfoCacheKey = CacheConstants.GatedCacheKeyPrefix + "appinstall_authenticated.{0}.{1}.{2}.{3}";
                CacheConstants.EntitlementsKey = "entitle.{0}.{1}";
                return CacheConstants
            }();
        WACUtils.CacheConstants = CacheConstants
    })(WACUtils = OfficeExt.WACUtils || (OfficeExt.WACUtils = {}))
})(OfficeExt || (OfficeExt = {}));
OSF.OUtil.setNamespace("Microsoft",window);
OSF.OUtil.setNamespace("Office",Microsoft);
OSF.OUtil.setNamespace("Common",Microsoft.Office);
Microsoft.Office.Common.InvokeType = {
    async: 0,
    sync: 1,
    asyncRegisterEvent: 2,
    asyncUnregisterEvent: 3,
    syncRegisterEvent: 4,
    syncUnregisterEvent: 5
};
OSF.SerializerVersion = {
    MsAjax: 0,
    Browser: 1
};
var OfficeExt;
(function(OfficeExt)
{
    function appSpecificCheckOriginFunction(allowed_domains, eventObj, origin, checkOriginFunction)
    {
        return false
    }
    OfficeExt.appSpecificCheckOrigin = appSpecificCheckOriginFunction
})(OfficeExt || (OfficeExt = {}));
Microsoft.Office.Common.InvokeType = {
    async: 0,
    sync: 1,
    asyncRegisterEvent: 2,
    asyncUnregisterEvent: 3,
    syncRegisterEvent: 4,
    syncUnregisterEvent: 5
};
Microsoft.Office.Common.InvokeResultCode = {
    noError: 0,
    errorInRequest: -1,
    errorHandlingRequest: -2,
    errorInResponse: -3,
    errorHandlingResponse: -4,
    errorHandlingRequestAccessDenied: -5,
    errorHandlingMethodCallTimedout: -6
};
Microsoft.Office.Common.MessageType = {
    request: 0,
    response: 1
};
Microsoft.Office.Common.ActionType = {
    invoke: 0,
    registerEvent: 1,
    unregisterEvent: 2
};
Microsoft.Office.Common.ResponseType = {
    forCalling: 0,
    forEventing: 1
};
Microsoft.Office.Common.MethodObject = function Microsoft_Office_Common_MethodObject(method, invokeType, blockingOthers)
{
    this._method = method;
    this._invokeType = invokeType;
    this._blockingOthers = blockingOthers
};
Microsoft.Office.Common.MethodObject.prototype = {
    getMethod: function Microsoft_Office_Common_MethodObject$getMethod()
    {
        return this._method
    },
    getInvokeType: function Microsoft_Office_Common_MethodObject$getInvokeType()
    {
        return this._invokeType
    },
    getBlockingFlag: function Microsoft_Office_Common_MethodObject$getBlockingFlag()
    {
        return this._blockingOthers
    }
};
Microsoft.Office.Common.EventMethodObject = function Microsoft_Office_Common_EventMethodObject(registerMethodObject, unregisterMethodObject)
{
    this._registerMethodObject = registerMethodObject;
    this._unregisterMethodObject = unregisterMethodObject
};
Microsoft.Office.Common.EventMethodObject.prototype = {
    getRegisterMethodObject: function Microsoft_Office_Common_EventMethodObject$getRegisterMethodObject()
    {
        return this._registerMethodObject
    },
    getUnregisterMethodObject: function Microsoft_Office_Common_EventMethodObject$getUnregisterMethodObject()
    {
        return this._unregisterMethodObject
    }
};
Microsoft.Office.Common.ServiceEndPoint = function Microsoft_Office_Common_ServiceEndPoint(serviceEndPointId)
{
    var e = Function._validateParams(arguments,[{
                name: "serviceEndPointId",
                type: String,
                mayBeNull: false
            }]);
    if(e)
        throw e;
    this._methodObjectList = {};
    this._eventHandlerProxyList = {};
    this._Id = serviceEndPointId;
    this._conversations = {};
    this._policyManager = null;
    this._appDomains = {};
    this._onHandleRequestError = null;
    this._addInSourceLocationSubdomainAllowedIsEnabled = false
};
Microsoft.Office.Common.ServiceEndPoint.prototype = {
    registerMethod: function Microsoft_Office_Common_ServiceEndPoint$registerMethod(methodName, method, invokeType, blockingOthers)
    {
        var e = Function._validateParams(arguments,[{
                    name: "methodName",
                    type: String,
                    mayBeNull: false
                },{
                    name: "method",
                    type: Function,
                    mayBeNull: false
                },{
                    name: "invokeType",
                    type: Number,
                    mayBeNull: false
                },{
                    name: "blockingOthers",
                    type: Boolean,
                    mayBeNull: false
                }]);
        if(e)
            throw e;
        if(invokeType !== Microsoft.Office.Common.InvokeType.async && invokeType !== Microsoft.Office.Common.InvokeType.sync)
            throw OsfMsAjaxFactory.msAjaxError.argument("invokeType");
        var methodObject = new Microsoft.Office.Common.MethodObject(method,invokeType,blockingOthers);
        this._methodObjectList[methodName] = methodObject
    },
    unregisterMethod: function Microsoft_Office_Common_ServiceEndPoint$unregisterMethod(methodName)
    {
        var e = Function._validateParams(arguments,[{
                    name: "methodName",
                    type: String,
                    mayBeNull: false
                }]);
        if(e)
            throw e;
        delete this._methodObjectList[methodName]
    },
    registerEvent: function Microsoft_Office_Common_ServiceEndPoint$registerEvent(eventName, registerMethod, unregisterMethod)
    {
        var e = Function._validateParams(arguments,[{
                    name: "eventName",
                    type: String,
                    mayBeNull: false
                },{
                    name: "registerMethod",
                    type: Function,
                    mayBeNull: false
                },{
                    name: "unregisterMethod",
                    type: Function,
                    mayBeNull: false
                }]);
        if(e)
            throw e;
        var methodObject = new Microsoft.Office.Common.EventMethodObject(new Microsoft.Office.Common.MethodObject(registerMethod,Microsoft.Office.Common.InvokeType.syncRegisterEvent,false),new Microsoft.Office.Common.MethodObject(unregisterMethod,Microsoft.Office.Common.InvokeType.syncUnregisterEvent,false));
        this._methodObjectList[eventName] = methodObject
    },
    registerEventEx: function Microsoft_Office_Common_ServiceEndPoint$registerEventEx(eventName, registerMethod, registerMethodInvokeType, unregisterMethod, unregisterMethodInvokeType)
    {
        var e = Function._validateParams(arguments,[{
                    name: "eventName",
                    type: String,
                    mayBeNull: false
                },{
                    name: "registerMethod",
                    type: Function,
                    mayBeNull: false
                },{
                    name: "registerMethodInvokeType",
                    type: Number,
                    mayBeNull: false
                },{
                    name: "unregisterMethod",
                    type: Function,
                    mayBeNull: false
                },{
                    name: "unregisterMethodInvokeType",
                    type: Number,
                    mayBeNull: false
                }]);
        if(e)
            throw e;
        var methodObject = new Microsoft.Office.Common.EventMethodObject(new Microsoft.Office.Common.MethodObject(registerMethod,registerMethodInvokeType,false),new Microsoft.Office.Common.MethodObject(unregisterMethod,unregisterMethodInvokeType,false));
        this._methodObjectList[eventName] = methodObject
    },
    unregisterEvent: function(eventName)
    {
        var e = Function._validateParams(arguments,[{
                    name: "eventName",
                    type: String,
                    mayBeNull: false
                }]);
        if(e)
            throw e;
        this.unregisterMethod(eventName)
    },
    registerConversation: function Microsoft_Office_Common_ServiceEndPoint$registerConversation(conversationId, conversationUrl, appDomains, serializerVersion, addInSourceLocationSubdomainAllowedIsEnabled)
    {
        var e = Function._validateParams(arguments,[{
                    name: "conversationId",
                    type: String,
                    mayBeNull: false
                },{
                    name: "conversationUrl",
                    type: String,
                    mayBeNull: false,
                    optional: true
                },{
                    name: "appDomains",
                    type: Object,
                    mayBeNull: true,
                    optional: true
                },{
                    name: "serializerVersion",
                    type: Number,
                    mayBeNull: true,
                    optional: true
                },{
                    name: "addInSourceLocationSubdomainAllowedIsEnabled",
                    type: Boolean,
                    mayBeNull: true,
                    optional: true
                }]);
        if(e)
            throw e;
        if(addInSourceLocationSubdomainAllowedIsEnabled)
            this._addInSourceLocationSubdomainAllowedIsEnabled = addInSourceLocationSubdomainAllowedIsEnabled;
        if(appDomains)
        {
            if(!(appDomains instanceof Array))
                throw OsfMsAjaxFactory.msAjaxError.argument("appDomains");
            this._appDomains[conversationId] = appDomains
        }
        this._conversations[conversationId] = {
            url: conversationUrl,
            serializerVersion: serializerVersion
        }
    },
    unregisterConversation: function Microsoft_Office_Common_ServiceEndPoint$unregisterConversation(conversationId)
    {
        var e = Function._validateParams(arguments,[{
                    name: "conversationId",
                    type: String,
                    mayBeNull: false
                }]);
        if(e)
            throw e;
        delete this._conversations[conversationId]
    },
    setPolicyManager: function Microsoft_Office_Common_ServiceEndPoint$setPolicyManager(policyManager)
    {
        var e = Function._validateParams(arguments,[{
                    name: "policyManager",
                    type: Object,
                    mayBeNull: false
                }]);
        if(e)
            throw e;
        if(!policyManager.checkPermission)
            throw OsfMsAjaxFactory.msAjaxError.argument("policyManager");
        this._policyManager = policyManager
    },
    getPolicyManager: function Microsoft_Office_Common_ServiceEndPoint$getPolicyManager()
    {
        return this._policyManager
    },
    dispose: function Microsoft_Office_Common_ServiceEndPoint$dispose()
    {
        this._methodObjectList = null;
        this._eventHandlerProxyList = null;
        this._Id = null;
        this._conversations = null;
        this._policyManager = null;
        this._appDomains = null;
        this._onHandleRequestError = null
    }
};
Microsoft.Office.Common.ClientEndPoint = function Microsoft_Office_Common_ClientEndPoint(conversationId, targetWindow, targetUrl, serializerVersion)
{
    var e = Function._validateParams(arguments,[{
                name: "conversationId",
                type: String,
                mayBeNull: false
            },{
                name: "targetWindow",
                mayBeNull: false
            },{
                name: "targetUrl",
                type: String,
                mayBeNull: false
            },{
                name: "serializerVersion",
                type: Number,
                mayBeNull: true,
                optional: true
            }]);
    if(e)
        throw e;
    try
    {
        if(!targetWindow.postMessage)
            throw OsfMsAjaxFactory.msAjaxError.argument("targetWindow");
    }
    catch(ex)
    {
        if(!Object.prototype.hasOwnProperty.call(targetWindow,"postMessage"))
            throw OsfMsAjaxFactory.msAjaxError.argument("targetWindow");
    }
    this._conversationId = conversationId;
    this._targetWindow = targetWindow;
    this._targetUrl = targetUrl;
    this._callingIndex = 0;
    this._callbackList = {};
    this._eventHandlerList = {};
    if(serializerVersion != null)
        this._serializerVersion = serializerVersion;
    else
        this._serializerVersion = OSF.SerializerVersion.Browser
};
Microsoft.Office.Common.ClientEndPoint.prototype = {
    invoke: function Microsoft_Office_Common_ClientEndPoint$invoke(targetMethodName, callback, param)
    {
        var e = Function._validateParams(arguments,[{
                    name: "targetMethodName",
                    type: String,
                    mayBeNull: false
                },{
                    name: "callback",
                    type: Function,
                    mayBeNull: true
                },{
                    name: "param",
                    mayBeNull: true
                }]);
        if(e)
            throw e;
        var correlationId = this._callingIndex++;
        var now = new Date;
        var callbackEntry = {
                callback: callback,
                createdOn: now.getTime()
            };
        if(param && typeof param === "object" && typeof param.__timeout__ === "number")
        {
            callbackEntry.timeout = param.__timeout__;
            delete param.__timeout__
        }
        this._callbackList[correlationId] = callbackEntry;
        try
        {
            var callRequest = new Microsoft.Office.Common.Request(targetMethodName,Microsoft.Office.Common.ActionType.invoke,this._conversationId,correlationId,param);
            var msg = Microsoft.Office.Common.MessagePackager.envelope(callRequest,this._serializerVersion);
            this._targetWindow.postMessage(msg,this._targetUrl);
            Microsoft.Office.Common.XdmCommunicationManager._startMethodTimeoutTimer()
        }
        catch(ex)
        {
            try
            {
                if(callback !== null)
                    callback(Microsoft.Office.Common.InvokeResultCode.errorInRequest,ex)
            }
            finally
            {
                delete this._callbackList[correlationId]
            }
        }
    },
    registerForEvent: function Microsoft_Office_Common_ClientEndPoint$registerForEvent(targetEventName, eventHandler, callback, data)
    {
        var e = Function._validateParams(arguments,[{
                    name: "targetEventName",
                    type: String,
                    mayBeNull: false
                },{
                    name: "eventHandler",
                    type: Function,
                    mayBeNull: false
                },{
                    name: "callback",
                    type: Function,
                    mayBeNull: true
                },{
                    name: "data",
                    mayBeNull: true,
                    optional: true
                }]);
        if(e)
            throw e;
        var correlationId = this._callingIndex++;
        var now = new Date;
        this._callbackList[correlationId] = {
            callback: callback,
            createdOn: now.getTime()
        };
        try
        {
            var callRequest = new Microsoft.Office.Common.Request(targetEventName,Microsoft.Office.Common.ActionType.registerEvent,this._conversationId,correlationId,data);
            var msg = Microsoft.Office.Common.MessagePackager.envelope(callRequest,this._serializerVersion);
            this._targetWindow.postMessage(msg,this._targetUrl);
            Microsoft.Office.Common.XdmCommunicationManager._startMethodTimeoutTimer();
            this._eventHandlerList[targetEventName] = eventHandler
        }
        catch(ex)
        {
            try
            {
                if(callback !== null)
                    callback(Microsoft.Office.Common.InvokeResultCode.errorInRequest,ex)
            }
            finally
            {
                delete this._callbackList[correlationId]
            }
        }
    },
    unregisterForEvent: function Microsoft_Office_Common_ClientEndPoint$unregisterForEvent(targetEventName, callback, data)
    {
        var e = Function._validateParams(arguments,[{
                    name: "targetEventName",
                    type: String,
                    mayBeNull: false
                },{
                    name: "callback",
                    type: Function,
                    mayBeNull: true
                },{
                    name: "data",
                    mayBeNull: true,
                    optional: true
                }]);
        if(e)
            throw e;
        var correlationId = this._callingIndex++;
        var now = new Date;
        this._callbackList[correlationId] = {
            callback: callback,
            createdOn: now.getTime()
        };
        try
        {
            var callRequest = new Microsoft.Office.Common.Request(targetEventName,Microsoft.Office.Common.ActionType.unregisterEvent,this._conversationId,correlationId,data);
            var msg = Microsoft.Office.Common.MessagePackager.envelope(callRequest,this._serializerVersion);
            this._targetWindow.postMessage(msg,this._targetUrl);
            Microsoft.Office.Common.XdmCommunicationManager._startMethodTimeoutTimer()
        }
        catch(ex)
        {
            try
            {
                if(callback !== null)
                    callback(Microsoft.Office.Common.InvokeResultCode.errorInRequest,ex)
            }
            finally
            {
                delete this._callbackList[correlationId]
            }
        }
        finally
        {
            delete this._eventHandlerList[targetEventName]
        }
    }
};
Microsoft.Office.Common.XdmCommunicationManager = function()
{
    var _invokerQueue = [];
    var _lastMessageProcessTime = null;
    var _messageProcessingTimer = null;
    var _processInterval = 10;
    var _blockingFlag = false;
    var _methodTimeoutTimer = null;
    var _methodTimeoutProcessInterval = 2e3;
    var _methodTimeoutDefault = 65e3;
    var _methodTimeout = _methodTimeoutDefault;
    var _serviceEndPoints = {};
    var _clientEndPoints = {};
    var _initialized = false;
    function _lookupServiceEndPoint(conversationId)
    {
        for(var id in _serviceEndPoints)
            if(_serviceEndPoints[id]._conversations[conversationId])
                return _serviceEndPoints[id];
        OsfMsAjaxFactory.msAjaxDebug.trace("Unknown conversation Id.");
        throw OsfMsAjaxFactory.msAjaxError.argument("conversationId");
    }
    function _lookupClientEndPoint(conversationId)
    {
        var clientEndPoint = _clientEndPoints[conversationId];
        if(!clientEndPoint)
            OsfMsAjaxFactory.msAjaxDebug.trace("Unknown conversation Id.");
        return clientEndPoint
    }
    function _lookupMethodObject(serviceEndPoint, messageObject)
    {
        var methodOrEventMethodObject = serviceEndPoint._methodObjectList[messageObject._actionName];
        if(!methodOrEventMethodObject)
        {
            OsfMsAjaxFactory.msAjaxDebug.trace("The specified method is not registered on service endpoint:" + messageObject._actionName);
            throw OsfMsAjaxFactory.msAjaxError.argument("messageObject");
        }
        var methodObject = null;
        if(messageObject._actionType === Microsoft.Office.Common.ActionType.invoke)
            methodObject = methodOrEventMethodObject;
        else if(messageObject._actionType === Microsoft.Office.Common.ActionType.registerEvent)
            methodObject = methodOrEventMethodObject.getRegisterMethodObject();
        else
            methodObject = methodOrEventMethodObject.getUnregisterMethodObject();
        return methodObject
    }
    function _enqueInvoker(invoker)
    {
        _invokerQueue.push(invoker)
    }
    function _dequeInvoker()
    {
        if(_messageProcessingTimer !== null)
        {
            if(!_blockingFlag)
                if(_invokerQueue.length > 0)
                {
                    var invoker = _invokerQueue.shift();
                    _executeCommand(invoker)
                }
                else
                {
                    clearInterval(_messageProcessingTimer);
                    _messageProcessingTimer = null
                }
        }
        else
            OsfMsAjaxFactory.msAjaxDebug.trace("channel is not ready.")
    }
    function _executeCommand(invoker)
    {
        _blockingFlag = invoker.getInvokeBlockingFlag();
        invoker.invoke();
        _lastMessageProcessTime = (new Date).getTime()
    }
    function _checkMethodTimeout()
    {
        if(_methodTimeoutTimer)
        {
            var clientEndPoint;
            var methodCallsNotTimedout = 0;
            var now = new Date;
            var timeoutValue;
            for(var conversationId in _clientEndPoints)
            {
                clientEndPoint = _clientEndPoints[conversationId];
                for(var correlationId in clientEndPoint._callbackList)
                {
                    var callbackEntry = clientEndPoint._callbackList[correlationId];
                    timeoutValue = callbackEntry.timeout ? callbackEntry.timeout : _methodTimeout;
                    if(timeoutValue >= 0 && Math.abs(now.getTime() - callbackEntry.createdOn) >= timeoutValue)
                        try
                        {
                            if(callbackEntry.callback)
                                callbackEntry.callback(Microsoft.Office.Common.InvokeResultCode.errorHandlingMethodCallTimedout,null)
                        }
                        finally
                        {
                            delete clientEndPoint._callbackList[correlationId]
                        }
                    else
                        methodCallsNotTimedout++
                }
            }
            if(methodCallsNotTimedout === 0)
            {
                clearInterval(_methodTimeoutTimer);
                _methodTimeoutTimer = null
            }
        }
        else
            OsfMsAjaxFactory.msAjaxDebug.trace("channel is not ready.")
    }
    function _postCallbackHandler()
    {
        _blockingFlag = false
    }
    function _registerListener(listener)
    {
        if(window.addEventListener)
            window.addEventListener("message",listener,false);
        else if(navigator.userAgent.indexOf("MSIE") > -1 && window.attachEvent)
            window.attachEvent("onmessage",listener);
        else
        {
            OsfMsAjaxFactory.msAjaxDebug.trace("Browser doesn't support the required API.");
            throw OsfMsAjaxFactory.msAjaxError.argument("Browser");
        }
    }
    function _checkOrigin(url, origin)
    {
        var res = false;
        if(url === true)
            return true;
        if(!url || !origin || !url.length || !origin.length)
            return res;
        var url_parser,
            org_parser;
        url_parser = document.createElement("a");
        org_parser = document.createElement("a");
        url_parser.href = url;
        org_parser.href = origin;
        res = _urlCompare(url_parser,org_parser);
        delete url_parser,org_parser;
        return res
    }
    function _checkOriginWithAppDomains(allowed_domains, origin)
    {
        var res = false;
        if(!origin || !origin.length || !allowed_domains || !(allowed_domains instanceof Array) || !allowed_domains.length)
            return res;
        var org_parser = document.createElement("a");
        var app_domain_parser = document.createElement("a");
        org_parser.href = origin;
        for(var i = 0; i < allowed_domains.length && !res; i++)
            if(allowed_domains[i].indexOf("://") !== -1)
            {
                app_domain_parser.href = allowed_domains[i];
                res = _urlCompare(org_parser,app_domain_parser)
            }
        delete org_parser,app_domain_parser;
        return res
    }
    function IsOriginSubdomainOfSourceLocation(sourceLocation, messageOrigin)
    {
        if(!sourceLocation || !messageOrigin)
            return false;
        var sourceLocationParser = document.createElement("a");
        sourceLocationParser.href = sourceLocation;
        var messageOriginParser = document.createElement("a");
        messageOriginParser.href = messageOrigin;
        var isSameProtocol = sourceLocationParser.protocol === messageOriginParser.protocol;
        var isSamePort = sourceLocationParser.port === messageOriginParser.port;
        var originHostName = messageOriginParser.hostname;
        var sourceLocationHostName = sourceLocationParser.hostname;
        var isSameDomain = originHostName === sourceLocationHostName;
        var isSubDomain = false;
        if(!isSameDomain && originHostName.length > sourceLocationHostName.length + 1)
            isSubDomain = originHostName.slice(-(sourceLocationHostName.length + 1)) === "." + sourceLocationHostName;
        var isSameDomainOrSubdomain = isSameDomain || isSubDomain;
        return isSamePort && isSameProtocol && isSameDomainOrSubdomain
    }
    function _urlCompare(url_parser1, url_parser2)
    {
        return url_parser1.hostname == url_parser2.hostname && url_parser1.protocol == url_parser2.protocol && url_parser1.port == url_parser2.port
    }
    function _receive(e)
    {
        if(!OSF)
            return;
        if(e.data != "")
        {
            var messageObject;
            var serializerVersion = OSF.SerializerVersion.Browser;
            var serializedMessage = e.data;
            try
            {
                messageObject = Microsoft.Office.Common.MessagePackager.unenvelope(serializedMessage,OSF.SerializerVersion.Browser);
                serializerVersion = messageObject._serializerVersion != null ? messageObject._serializerVersion : serializerVersion
            }
            catch(ex)
            {
                return
            }
            if(messageObject._messageType === Microsoft.Office.Common.MessageType.request)
            {
                var requesterUrl = e.origin == null || e.origin == "null" ? messageObject._origin : e.origin;
                try
                {
                    var serviceEndPoint = _lookupServiceEndPoint(messageObject._conversationId);
                    var conversation = serviceEndPoint._conversations[messageObject._conversationId];
                    serializerVersion = conversation.serializerVersion != null ? conversation.serializerVersion : serializerVersion;
                    var allowedDomains = [conversation.url].concat(serviceEndPoint._appDomains[messageObject._conversationId]);
                    if(!_checkOriginWithAppDomains(allowedDomains,e.origin))
                        if(!OfficeExt.appSpecificCheckOrigin(allowedDomains,e,messageObject._origin,_checkOriginWithAppDomains))
                        {
                            var isOriginSubdomain = serviceEndPoint._addInSourceLocationSubdomainAllowedIsEnabled && IsOriginSubdomainOfSourceLocation(conversation.url,e.origin);
                            if(!isOriginSubdomain)
                                throw"Failed origin check";
                        }
                    var policyManager = serviceEndPoint.getPolicyManager();
                    if(policyManager && !policyManager.checkPermission(messageObject._conversationId,messageObject._actionName,messageObject._data))
                        throw"Access Denied";
                    var methodObject = _lookupMethodObject(serviceEndPoint,messageObject);
                    var invokeCompleteCallback = new Microsoft.Office.Common.InvokeCompleteCallback(e.source,requesterUrl,messageObject._actionName,messageObject._conversationId,messageObject._correlationId,_postCallbackHandler,serializerVersion);
                    var invoker = new Microsoft.Office.Common.Invoker(methodObject,messageObject._data,invokeCompleteCallback,serviceEndPoint._eventHandlerProxyList,messageObject._conversationId,messageObject._actionName,serializerVersion);
                    var shouldEnque = true;
                    if(_messageProcessingTimer == null)
                        if((_lastMessageProcessTime == null || (new Date).getTime() - _lastMessageProcessTime > _processInterval) && !_blockingFlag)
                        {
                            _executeCommand(invoker);
                            shouldEnque = false
                        }
                        else
                            _messageProcessingTimer = setInterval(_dequeInvoker,_processInterval);
                    if(shouldEnque)
                        _enqueInvoker(invoker)
                }
                catch(ex)
                {
                    if(serviceEndPoint && serviceEndPoint._onHandleRequestError)
                        serviceEndPoint._onHandleRequestError(messageObject,ex);
                    var errorCode = Microsoft.Office.Common.InvokeResultCode.errorHandlingRequest;
                    if(ex == "Access Denied")
                        errorCode = Microsoft.Office.Common.InvokeResultCode.errorHandlingRequestAccessDenied;
                    var callResponse = new Microsoft.Office.Common.Response(messageObject._actionName,messageObject._conversationId,messageObject._correlationId,errorCode,Microsoft.Office.Common.ResponseType.forCalling,ex);
                    var envelopedResult = Microsoft.Office.Common.MessagePackager.envelope(callResponse,serializerVersion);
                    var canPostMessage = false;
                    try
                    {
                        canPostMessage = !!(e.source && e.source.postMessage)
                    }
                    catch(ex){}
                    if(canPostMessage)
                        e.source.postMessage(envelopedResult,requesterUrl)
                }
            }
            else if(messageObject._messageType === Microsoft.Office.Common.MessageType.response)
            {
                var clientEndPoint = _lookupClientEndPoint(messageObject._conversationId);
                if(!clientEndPoint)
                    return;
                clientEndPoint._serializerVersion = serializerVersion;
                if(!_checkOrigin(clientEndPoint._targetUrl,e.origin))
                    throw"Failed orgin check";
                if(messageObject._responseType === Microsoft.Office.Common.ResponseType.forCalling)
                {
                    var callbackEntry = clientEndPoint._callbackList[messageObject._correlationId];
                    if(callbackEntry)
                        try
                        {
                            if(callbackEntry.callback)
                                callbackEntry.callback(messageObject._errorCode,messageObject._data)
                        }
                        finally
                        {
                            delete clientEndPoint._callbackList[messageObject._correlationId]
                        }
                }
                else
                {
                    var eventhandler = clientEndPoint._eventHandlerList[messageObject._actionName];
                    if(eventhandler !== undefined && eventhandler !== null)
                        eventhandler(messageObject._data)
                }
            }
            else
                return
        }
    }
    function _initialize()
    {
        if(!_initialized)
        {
            _registerListener(_receive);
            _initialized = true
        }
    }
    return{
            connect: function Microsoft_Office_Common_XdmCommunicationManager$connect(conversationId, targetWindow, targetUrl, serializerVersion)
            {
                var clientEndPoint = _clientEndPoints[conversationId];
                if(!clientEndPoint)
                {
                    _initialize();
                    clientEndPoint = new Microsoft.Office.Common.ClientEndPoint(conversationId,targetWindow,targetUrl,serializerVersion);
                    _clientEndPoints[conversationId] = clientEndPoint
                }
                return clientEndPoint
            },
            getClientEndPoint: function Microsoft_Office_Common_XdmCommunicationManager$getClientEndPoint(conversationId)
            {
                var e = Function._validateParams(arguments,[{
                            name: "conversationId",
                            type: String,
                            mayBeNull: false
                        }]);
                if(e)
                    throw e;
                return _clientEndPoints[conversationId]
            },
            createServiceEndPoint: function Microsoft_Office_Common_XdmCommunicationManager$createServiceEndPoint(serviceEndPointId)
            {
                _initialize();
                var serviceEndPoint = new Microsoft.Office.Common.ServiceEndPoint(serviceEndPointId);
                _serviceEndPoints[serviceEndPointId] = serviceEndPoint;
                return serviceEndPoint
            },
            getServiceEndPoint: function Microsoft_Office_Common_XdmCommunicationManager$getServiceEndPoint(serviceEndPointId)
            {
                var e = Function._validateParams(arguments,[{
                            name: "serviceEndPointId",
                            type: String,
                            mayBeNull: false
                        }]);
                if(e)
                    throw e;
                return _serviceEndPoints[serviceEndPointId]
            },
            deleteClientEndPoint: function Microsoft_Office_Common_XdmCommunicationManager$deleteClientEndPoint(conversationId)
            {
                var e = Function._validateParams(arguments,[{
                            name: "conversationId",
                            type: String,
                            mayBeNull: false
                        }]);
                if(e)
                    throw e;
                delete _clientEndPoints[conversationId]
            },
            deleteServiceEndPoint: function Microsoft_Office_Common_XdmCommunicationManager$deleteServiceEndPoint(serviceEndPointId)
            {
                var e = Function._validateParams(arguments,[{
                            name: "serviceEndPointId",
                            type: String,
                            mayBeNull: false
                        }]);
                if(e)
                    throw e;
                delete _serviceEndPoints[serviceEndPointId]
            },
            checkUrlWithAppDomains: function Microsoft_Office_Common_XdmCommunicationManager$_checkUrlWithAppDomains(appDomains, origin)
            {
                return _checkOriginWithAppDomains(appDomains,origin)
            },
            _setMethodTimeout: function Microsoft_Office_Common_XdmCommunicationManager$_setMethodTimeout(methodTimeout)
            {
                var e = Function._validateParams(arguments,[{
                            name: "methodTimeout",
                            type: Number,
                            mayBeNull: false
                        }]);
                if(e)
                    throw e;
                _methodTimeout = methodTimeout <= 0 ? _methodTimeoutDefault : methodTimeout
            },
            _startMethodTimeoutTimer: function Microsoft_Office_Common_XdmCommunicationManager$_startMethodTimeoutTimer()
            {
                if(!_methodTimeoutTimer)
                    _methodTimeoutTimer = setInterval(_checkMethodTimeout,_methodTimeoutProcessInterval)
            }
        }
}();
Microsoft.Office.Common.Message = function Microsoft_Office_Common_Message(messageType, actionName, conversationId, correlationId, data)
{
    var e = Function._validateParams(arguments,[{
                name: "messageType",
                type: Number,
                mayBeNull: false
            },{
                name: "actionName",
                type: String,
                mayBeNull: false
            },{
                name: "conversationId",
                type: String,
                mayBeNull: false
            },{
                name: "correlationId",
                mayBeNull: false
            },{
                name: "data",
                mayBeNull: true,
                optional: true
            }]);
    if(e)
        throw e;
    this._messageType = messageType;
    this._actionName = actionName;
    this._conversationId = conversationId;
    this._correlationId = correlationId;
    this._origin = window.location.href;
    if(typeof data == "undefined")
        this._data = null;
    else
        this._data = data
};
Microsoft.Office.Common.Message.prototype = {
    getActionName: function Microsoft_Office_Common_Message$getActionName()
    {
        return this._actionName
    },
    getConversationId: function Microsoft_Office_Common_Message$getConversationId()
    {
        return this._conversationId
    },
    getCorrelationId: function Microsoft_Office_Common_Message$getCorrelationId()
    {
        return this._correlationId
    },
    getOrigin: function Microsoft_Office_Common_Message$getOrigin()
    {
        return this._origin
    },
    getData: function Microsoft_Office_Common_Message$getData()
    {
        return this._data
    },
    getMessageType: function Microsoft_Office_Common_Message$getMessageType()
    {
        return this._messageType
    }
};
Microsoft.Office.Common.Request = function Microsoft_Office_Common_Request(actionName, actionType, conversationId, correlationId, data)
{
    Microsoft.Office.Common.Request.uber.constructor.call(this,Microsoft.Office.Common.MessageType.request,actionName,conversationId,correlationId,data);
    this._actionType = actionType
};
OSF.OUtil.extend(Microsoft.Office.Common.Request,Microsoft.Office.Common.Message);
Microsoft.Office.Common.Request.prototype.getActionType = function Microsoft_Office_Common_Request$getActionType()
{
    return this._actionType
};
Microsoft.Office.Common.Response = function Microsoft_Office_Common_Response(actionName, conversationId, correlationId, errorCode, responseType, data)
{
    Microsoft.Office.Common.Response.uber.constructor.call(this,Microsoft.Office.Common.MessageType.response,actionName,conversationId,correlationId,data);
    this._errorCode = errorCode;
    this._responseType = responseType
};
OSF.OUtil.extend(Microsoft.Office.Common.Response,Microsoft.Office.Common.Message);
Microsoft.Office.Common.Response.prototype.getErrorCode = function Microsoft_Office_Common_Response$getErrorCode()
{
    return this._errorCode
};
Microsoft.Office.Common.Response.prototype.getResponseType = function Microsoft_Office_Common_Response$getResponseType()
{
    return this._responseType
};
Microsoft.Office.Common.MessagePackager = {
    envelope: function Microsoft_Office_Common_MessagePackager$envelope(messageObject, serializerVersion)
    {
        if(typeof messageObject === "object")
            messageObject._serializerVersion = OSF.SerializerVersion.Browser;
        return JSON.stringify(messageObject)
    },
    unenvelope: function Microsoft_Office_Common_MessagePackager$unenvelope(messageObject, serializerVersion)
    {
        return JSON.parse(messageObject)
    }
};
Microsoft.Office.Common.ResponseSender = function Microsoft_Office_Common_ResponseSender(requesterWindow, requesterUrl, actionName, conversationId, correlationId, responseType, serializerVersion)
{
    var e = Function._validateParams(arguments,[{
                name: "requesterWindow",
                mayBeNull: false
            },{
                name: "requesterUrl",
                type: String,
                mayBeNull: false
            },{
                name: "actionName",
                type: String,
                mayBeNull: false
            },{
                name: "conversationId",
                type: String,
                mayBeNull: false
            },{
                name: "correlationId",
                mayBeNull: false
            },{
                name: "responsetype",
                type: Number,
                maybeNull: false
            },{
                name: "serializerVersion",
                type: Number,
                maybeNull: true,
                optional: true
            }]);
    if(e)
        throw e;
    this._requesterWindow = requesterWindow;
    this._requesterUrl = requesterUrl;
    this._actionName = actionName;
    this._conversationId = conversationId;
    this._correlationId = correlationId;
    this._invokeResultCode = Microsoft.Office.Common.InvokeResultCode.noError;
    this._responseType = responseType;
    var me = this;
    this._send = function(result)
    {
        try
        {
            var response = new Microsoft.Office.Common.Response(me._actionName,me._conversationId,me._correlationId,me._invokeResultCode,me._responseType,result);
            var envelopedResult = Microsoft.Office.Common.MessagePackager.envelope(response,serializerVersion);
            me._requesterWindow.postMessage(envelopedResult,me._requesterUrl)
        }
        catch(ex)
        {
            OsfMsAjaxFactory.msAjaxDebug.trace("ResponseSender._send error:" + ex.message)
        }
    }
};
Microsoft.Office.Common.ResponseSender.prototype = {
    getRequesterWindow: function Microsoft_Office_Common_ResponseSender$getRequesterWindow()
    {
        return this._requesterWindow
    },
    getRequesterUrl: function Microsoft_Office_Common_ResponseSender$getRequesterUrl()
    {
        return this._requesterUrl
    },
    getActionName: function Microsoft_Office_Common_ResponseSender$getActionName()
    {
        return this._actionName
    },
    getConversationId: function Microsoft_Office_Common_ResponseSender$getConversationId()
    {
        return this._conversationId
    },
    getCorrelationId: function Microsoft_Office_Common_ResponseSender$getCorrelationId()
    {
        return this._correlationId
    },
    getSend: function Microsoft_Office_Common_ResponseSender$getSend()
    {
        return this._send
    },
    setResultCode: function Microsoft_Office_Common_ResponseSender$setResultCode(resultCode)
    {
        this._invokeResultCode = resultCode
    }
};
Microsoft.Office.Common.InvokeCompleteCallback = function Microsoft_Office_Common_InvokeCompleteCallback(requesterWindow, requesterUrl, actionName, conversationId, correlationId, postCallbackHandler, serializerVersion)
{
    Microsoft.Office.Common.InvokeCompleteCallback.uber.constructor.call(this,requesterWindow,requesterUrl,actionName,conversationId,correlationId,Microsoft.Office.Common.ResponseType.forCalling,serializerVersion);
    this._postCallbackHandler = postCallbackHandler;
    var me = this;
    this._send = function(result, responseCode)
    {
        if(responseCode != undefined)
            me._invokeResultCode = responseCode;
        try
        {
            var response = new Microsoft.Office.Common.Response(me._actionName,me._conversationId,me._correlationId,me._invokeResultCode,me._responseType,result);
            var envelopedResult = Microsoft.Office.Common.MessagePackager.envelope(response,serializerVersion);
            me._requesterWindow.postMessage(envelopedResult,me._requesterUrl);
            me._postCallbackHandler()
        }
        catch(ex)
        {
            OsfMsAjaxFactory.msAjaxDebug.trace("InvokeCompleteCallback._send error:" + ex.message)
        }
    }
};
OSF.OUtil.extend(Microsoft.Office.Common.InvokeCompleteCallback,Microsoft.Office.Common.ResponseSender);
Microsoft.Office.Common.Invoker = function Microsoft_Office_Common_Invoker(methodObject, paramValue, invokeCompleteCallback, eventHandlerProxyList, conversationId, eventName, serializerVersion)
{
    var e = Function._validateParams(arguments,[{
                name: "methodObject",
                mayBeNull: false
            },{
                name: "paramValue",
                mayBeNull: true
            },{
                name: "invokeCompleteCallback",
                mayBeNull: false
            },{
                name: "eventHandlerProxyList",
                mayBeNull: true
            },{
                name: "conversationId",
                type: String,
                mayBeNull: false
            },{
                name: "eventName",
                type: String,
                mayBeNull: false
            },{
                name: "serializerVersion",
                type: Number,
                mayBeNull: true,
                optional: true
            }]);
    if(e)
        throw e;
    this._methodObject = methodObject;
    this._param = paramValue;
    this._invokeCompleteCallback = invokeCompleteCallback;
    this._eventHandlerProxyList = eventHandlerProxyList;
    this._conversationId = conversationId;
    this._eventName = eventName;
    this._serializerVersion = serializerVersion
};
Microsoft.Office.Common.Invoker.prototype = {
    invoke: function Microsoft_Office_Common_Invoker$invoke()
    {
        try
        {
            var result;
            switch(this._methodObject.getInvokeType())
            {
                case Microsoft.Office.Common.InvokeType.async:
                    this._methodObject.getMethod()(this._param,this._invokeCompleteCallback.getSend());
                    break;
                case Microsoft.Office.Common.InvokeType.sync:
                    result = this._methodObject.getMethod()(this._param);
                    this._invokeCompleteCallback.getSend()(result);
                    break;
                case Microsoft.Office.Common.InvokeType.syncRegisterEvent:
                    var eventHandlerProxy = this._createEventHandlerProxyObject(this._invokeCompleteCallback);
                    result = this._methodObject.getMethod()(eventHandlerProxy.getSend(),this._param);
                    this._eventHandlerProxyList[this._conversationId + this._eventName] = eventHandlerProxy.getSend();
                    this._invokeCompleteCallback.getSend()(result);
                    break;
                case Microsoft.Office.Common.InvokeType.syncUnregisterEvent:
                    var eventHandler = this._eventHandlerProxyList[this._conversationId + this._eventName];
                    result = this._methodObject.getMethod()(eventHandler,this._param);
                    delete this._eventHandlerProxyList[this._conversationId + this._eventName];
                    this._invokeCompleteCallback.getSend()(result);
                    break;
                case Microsoft.Office.Common.InvokeType.asyncRegisterEvent:
                    var eventHandlerProxyAsync = this._createEventHandlerProxyObject(this._invokeCompleteCallback);
                    this._methodObject.getMethod()(eventHandlerProxyAsync.getSend(),this._invokeCompleteCallback.getSend(),this._param);
                    this._eventHandlerProxyList[this._callerId + this._eventName] = eventHandlerProxyAsync.getSend();
                    break;
                case Microsoft.Office.Common.InvokeType.asyncUnregisterEvent:
                    var eventHandlerAsync = this._eventHandlerProxyList[this._callerId + this._eventName];
                    this._methodObject.getMethod()(eventHandlerAsync,this._invokeCompleteCallback.getSend(),this._param);
                    delete this._eventHandlerProxyList[this._callerId + this._eventName];
                    break;
                default:
                    break
            }
        }
        catch(ex)
        {
            this._invokeCompleteCallback.setResultCode(Microsoft.Office.Common.InvokeResultCode.errorInResponse);
            this._invokeCompleteCallback.getSend()(ex)
        }
    },
    getInvokeBlockingFlag: function Microsoft_Office_Common_Invoker$getInvokeBlockingFlag()
    {
        return this._methodObject.getBlockingFlag()
    },
    _createEventHandlerProxyObject: function Microsoft_Office_Common_Invoker$_createEventHandlerProxyObject(invokeCompleteObject)
    {
        return new Microsoft.Office.Common.ResponseSender(invokeCompleteObject.getRequesterWindow(),invokeCompleteObject.getRequesterUrl(),invokeCompleteObject.getActionName(),invokeCompleteObject.getConversationId(),invokeCompleteObject.getCorrelationId(),Microsoft.Office.Common.ResponseType.forEventing,this._serializerVersion)
    }
};
OSF.OUtil.setNamespace("WAC",OSF.DDA);
OSF.DDA.WAC.UniqueArguments = {
    Data: "Data",
    Properties: "Properties",
    BindingRequest: "DdaBindingsMethod",
    BindingResponse: "Bindings",
    SingleBindingResponse: "singleBindingResponse",
    GetData: "DdaGetBindingData",
    AddRowsColumns: "DdaAddRowsColumns",
    SetData: "DdaSetBindingData",
    ClearFormats: "DdaClearBindingFormats",
    SetFormats: "DdaSetBindingFormats",
    SettingsRequest: "DdaSettingsMethod",
    BindingEventSource: "ddaBinding",
    ArrayData: "ArrayData"
};
OSF.OUtil.setNamespace("Delegate",OSF.DDA.WAC);
OSF.DDA.WAC.Delegate.SpecialProcessor = function OSF_DDA_WAC_Delegate_SpecialProcessor()
{
    var complexTypes = [OSF.DDA.WAC.UniqueArguments.SingleBindingResponse,OSF.DDA.WAC.UniqueArguments.BindingRequest,OSF.DDA.WAC.UniqueArguments.BindingResponse,OSF.DDA.WAC.UniqueArguments.GetData,OSF.DDA.WAC.UniqueArguments.AddRowsColumns,OSF.DDA.WAC.UniqueArguments.SetData,OSF.DDA.WAC.UniqueArguments.ClearFormats,OSF.DDA.WAC.UniqueArguments.SetFormats,OSF.DDA.WAC.UniqueArguments.SettingsRequest,OSF.DDA.WAC.UniqueArguments.BindingEventSource];
    var dynamicTypes = {};
    OSF.DDA.WAC.Delegate.SpecialProcessor.uber.constructor.call(this,complexTypes,dynamicTypes)
};
OSF.OUtil.extend(OSF.DDA.WAC.Delegate.SpecialProcessor,OSF.DDA.SpecialProcessor);
OSF.DDA.WAC.Delegate.ParameterMap = OSF.DDA.getDecoratedParameterMap(new OSF.DDA.WAC.Delegate.SpecialProcessor,[]);
OSF.OUtil.setNamespace("WAC",OSF.DDA);
OSF.OUtil.setNamespace("Delegate",OSF.DDA.WAC);
OSF.DDA.WAC.getDelegateMethods = function OSF_DDA_WAC_getDelegateMethods()
{
    var delegateMethods = {};
    delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync] = OSF.DDA.WAC.Delegate.executeAsync;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync] = OSF.DDA.WAC.Delegate.registerEventAsync;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync] = OSF.DDA.WAC.Delegate.unregisterEventAsync;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog] = OSF.DDA.WAC.Delegate.openDialog;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.MessageParent] = OSF.DDA.WAC.Delegate.messageParent;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.SendMessage] = OSF.DDA.WAC.Delegate.sendMessage;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog] = OSF.DDA.WAC.Delegate.closeDialog;
    return delegateMethods
};
OSF.DDA.WAC.Delegate.version = 1;
OSF.DDA.WAC.Delegate.executeAsync = function OSF_DDA_WAC_Delegate$executeAsync(args)
{
    if(!args.hostCallArgs)
        args.hostCallArgs = {};
    args.hostCallArgs["DdaMethod"] = {
        ControlId: OSF._OfficeAppFactory.getId(),
        Version: OSF.DDA.WAC.Delegate.version,
        DispatchId: args.dispId
    };
    args.hostCallArgs["__timeout__"] = -1;
    if(args.onCalling)
        args.onCalling();
    if(!OSF.getClientEndPoint())
        return;
    OSF.getClientEndPoint().invoke("executeMethod",function OSF_DDA_WAC_Delegate$OMFacade$OnResponse(xdmStatus, payload)
    {
        if(args.onReceiving)
            args.onReceiving();
        var error;
        if(xdmStatus == Microsoft.Office.Common.InvokeResultCode.noError)
        {
            OSF.DDA.WAC.Delegate.version = payload["Version"];
            error = payload["Error"]
        }
        else
            switch(xdmStatus)
            {
                case Microsoft.Office.Common.InvokeResultCode.errorHandlingRequestAccessDenied:
                    error = OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
                    break;
                default:
                    error = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                    break
            }
        if(args.onComplete)
            args.onComplete(error,payload)
    },args.hostCallArgs)
};
OSF.DDA.WAC.Delegate._getOnAfterRegisterEvent = function OSF_DDA_WAC_Delegate$GetOnAfterRegisterEvent(register, args)
{
    var startTime = (new Date).getTime();
    return function OSF_DDA_WAC_Delegate$OnAfterRegisterEvent(xdmStatus, payload)
        {
            if(args.onReceiving)
                args.onReceiving();
            var status;
            if(xdmStatus != Microsoft.Office.Common.InvokeResultCode.noError)
                switch(xdmStatus)
                {
                    case Microsoft.Office.Common.InvokeResultCode.errorHandlingRequestAccessDenied:
                        status = OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
                        break;
                    default:
                        status = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                        break
                }
            else if(payload)
                if(payload["Error"])
                    status = payload["Error"];
                else
                    status = OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
            else
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
            if(args.onComplete)
                args.onComplete(status);
            if(OSF.AppTelemetry)
                OSF.AppTelemetry.onRegisterDone(register,args.dispId,Math.abs((new Date).getTime() - startTime),status)
        }
};
OSF.DDA.WAC.Delegate.registerEventAsync = function OSF_DDA_WAC_Delegate$RegisterEventAsync(args)
{
    if(args.onCalling)
        args.onCalling();
    if(!OSF.getClientEndPoint())
        return;
    OSF.getClientEndPoint().registerForEvent(OSF.DDA.getXdmEventName(args.targetId,args.eventType),function OSF_DDA_WACOMFacade$OnEvent(payload)
    {
        if(args.onEvent)
            args.onEvent(payload);
        if(OSF.AppTelemetry)
            OSF.AppTelemetry.onEventDone(args.dispId)
    },OSF.DDA.WAC.Delegate._getOnAfterRegisterEvent(true,args),{
        controlId: OSF._OfficeAppFactory.getId(),
        eventDispId: args.dispId,
        targetId: args.targetId
    })
};
OSF.DDA.WAC.Delegate.unregisterEventAsync = function OSF_DDA_WAC_Delegate$UnregisterEventAsync(args)
{
    if(args.onCalling)
        args.onCalling();
    if(!OSF.getClientEndPoint())
        return;
    OSF.getClientEndPoint().unregisterForEvent(OSF.DDA.getXdmEventName(args.targetId,args.eventType),OSF.DDA.WAC.Delegate._getOnAfterRegisterEvent(false,args),{
        controlId: OSF._OfficeAppFactory.getId(),
        eventDispId: args.dispId,
        targetId: args.targetId
    })
};
OSF.OUtil.setNamespace("WebApp",OSF);
OSF.WebApp.AddHostInfoAndXdmInfo = function OSF_WebApp$AddHostInfoAndXdmInfo(url)
{
    if(OSF._OfficeAppFactory.getWindowLocationSearch && OSF._OfficeAppFactory.getWindowLocationHash)
        return url + OSF._OfficeAppFactory.getWindowLocationSearch() + OSF._OfficeAppFactory.getWindowLocationHash();
    else
        return url
};
OSF.WebApp._UpdateLinksForHostAndXdmInfo = function OSF_WebApp$_UpdateLinksForHostAndXdmInfo()
{
    var links = document.querySelectorAll("a[data-officejs-navigate]");
    for(var i = 0; i < links.length; i++)
        if(OSF.WebApp._isGoodUrl(links[i].href))
            links[i].href = OSF.WebApp.AddHostInfoAndXdmInfo(links[i].href);
    var forms = document.querySelectorAll("form[data-officejs-navigate]");
    for(var i = 0; i < forms.length; i++)
    {
        var form = forms[i];
        if(OSF.WebApp._isGoodUrl(form.action))
            form.action = OSF.WebApp.AddHostInfoAndXdmInfo(form.action)
    }
};
OSF.WebApp._isGoodUrl = function OSF_WebApp$_isGoodUrl(url)
{
    if(typeof url == "undefined")
        return false;
    url = url.trim();
    var colonIndex = url.indexOf(":");
    var protocol = colonIndex > 0 ? url.substr(0,colonIndex) : null;
    var goodUrl = protocol !== null ? protocol.toLowerCase() === "http" || protocol.toLowerCase() === "https" : true;
    goodUrl = goodUrl && url != "#" && url != "/" && url != "" && url != OSF._OfficeAppFactory.getWebAppState().webAppUrl;
    return goodUrl
};
OSF.InitializationHelper = function OSF_InitializationHelper(hostInfo, webAppState, context, settings, hostFacade)
{
    this._hostInfo = hostInfo;
    this._webAppState = webAppState;
    this._context = context;
    this._settings = settings;
    this._hostFacade = hostFacade;
    this._appContext = {};
    this._tabbableElements = "a[href]:not([tabindex='-1'])," + "area[href]:not([tabindex='-1'])," + "button:not([disabled]):not([tabindex='-1'])," + "input:not([disabled]):not([tabindex='-1'])," + "select:not([disabled]):not([tabindex='-1'])," + "textarea:not([disabled]):not([tabindex='-1'])," + "*[tabindex]:not([tabindex='-1'])," + "*[contenteditable]:not([disabled]):not([tabindex='-1'])";
    this._initializeSettings = function OSF_InitializationHelper$initializeSettings(appContext, refreshSupported)
    {
        var settings;
        var serializedSettings = appContext.get_settings();
        var osfSessionStorage = OSF.OUtil.getSessionStorage();
        if(osfSessionStorage)
        {
            var storageSettings = osfSessionStorage.getItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey());
            if(storageSettings)
                serializedSettings = JSON.parse(storageSettings);
            else
            {
                storageSettings = JSON.stringify(serializedSettings);
                osfSessionStorage.setItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey(),storageSettings)
            }
        }
        var deserializedSettings = OSF.DDA.SettingsManager.deserializeSettings(serializedSettings);
        if(refreshSupported)
            settings = new OSF.DDA.RefreshableSettings(deserializedSettings);
        else
            settings = new OSF.DDA.Settings(deserializedSettings);
        return settings
    };
    var windowOpen = function OSF_InitializationHelper$windowOpen(windowObj)
        {
            var proxy = window.open;
            windowObj.open = function(strUrl, strWindowName, strWindowFeatures)
            {
                var windowObject = null;
                try
                {
                    windowObject = proxy(strUrl,strWindowName,strWindowFeatures)
                }
                catch(ex)
                {
                    if(OSF.AppTelemetry)
                        OSF.AppTelemetry.logAppCommonMessage("Exception happens at windowOpen." + ex)
                }
                if(!windowObject)
                {
                    var params = {
                            strUrl: strUrl,
                            strWindowName: strWindowName,
                            strWindowFeatures: strWindowFeatures
                        };
                    if(OSF._OfficeAppFactory.getClientEndPoint())
                        OSF._OfficeAppFactory.getClientEndPoint().invoke("ContextActivationManager_openWindowInHost",null,params)
                }
                return windowObject
            }
        };
    windowOpen(window)
};
OSF.InitializationHelper.prototype.saveAndSetDialogInfo = function OSF_InitializationHelper$saveAndSetDialogInfo(hostInfoValue)
{
    var getAppIdFromWindowLocation = function OSF_InitializationHelper$getAppIdFromWindowLocation()
        {
            var xdmInfoValue = OSF.OUtil.parseXdmInfo(true);
            if(xdmInfoValue)
            {
                var items = xdmInfoValue.split("|");
                return items[1]
            }
            return null
        };
    var osfSessionStorage = OSF.OUtil.getSessionStorage();
    if(osfSessionStorage)
    {
        if(!hostInfoValue)
            hostInfoValue = OSF.OUtil.parseHostInfoFromWindowName(true,OSF._OfficeAppFactory.getWindowName());
        if(hostInfoValue && hostInfoValue.indexOf("isDialog") > -1)
        {
            var appId = getAppIdFromWindowLocation();
            if(appId != null)
                osfSessionStorage.setItem(appId + "IsDialog","true");
            this._hostInfo.isDialog = true;
            return
        }
        this._hostInfo.isDialog = osfSessionStorage.getItem(OSF.OUtil.getXdmFieldValue(OSF.XdmFieldName.AppId,false) + "IsDialog") != null ? true : false
    }
};
OSF.InitializationHelper.prototype.getAppContext = function OSF_InitializationHelper$getAppContext(wnd, gotAppContext)
{
    var me = this;
    var getInvocationCallbackWebApp = function OSF_InitializationHelper_getAppContextAsync$getInvocationCallbackWebApp(errorCode, appContext)
        {
            var settings;
            if(appContext._appName === OSF.AppName.ExcelWebApp)
            {
                var serializedSettings = appContext._settings;
                settings = {};
                for(var index in serializedSettings)
                {
                    var setting = serializedSettings[index];
                    settings[setting[0]] = setting[1]
                }
            }
            else
                settings = appContext._settings;
            if(appContext._appName === OSF.AppName.OutlookWebApp && !!appContext._requirementMatrix && appContext._requirementMatrix.indexOf("react") == -1)
                OSF.AgaveHostAction.SendTelemetryEvent = undefined;
            if(!me._hostInfo.isDialog || window.opener == null)
            {
                var pageUrl = window.location.href;
                me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost",null,[me._webAppState.id,OSF.AgaveHostAction.UpdateTargetUrl,pageUrl])
            }
            if(errorCode === 0 && appContext._id != undefined && appContext._appName != undefined && appContext._appVersion != undefined && appContext._appUILocale != undefined && appContext._dataLocale != undefined && appContext._docUrl != undefined && appContext._clientMode != undefined && appContext._settings != undefined && appContext._reason != undefined)
            {
                me._appContext = appContext;
                var appInstanceId = appContext._appInstanceId ? appContext._appInstanceId : appContext._id;
                var touchEnabled = false;
                var commerceAllowed = true;
                var minorVersion = 0;
                if(appContext._appMinorVersion != undefined)
                    minorVersion = appContext._appMinorVersion;
                var requirementMatrix = undefined;
                if(appContext._requirementMatrix != undefined)
                    requirementMatrix = appContext._requirementMatrix;
                appContext.eToken = appContext.eToken ? appContext.eToken : "";
                var returnedContext = new OSF.OfficeAppContext(appContext._id,appContext._appName,appContext._appVersion,appContext._appUILocale,appContext._dataLocale,appContext._docUrl,appContext._clientMode,settings,appContext._reason,appContext._osfControlType,appContext._eToken,appContext._correlationId,appInstanceId,touchEnabled,commerceAllowed,minorVersion,requirementMatrix,appContext._hostCustomMessage,appContext._hostFullVersion,appContext._clientWindowHeight,appContext._clientWindowWidth,appContext._addinName,appContext._appDomains,appContext._dialogRequirementMatrix);
                returnedContext._wacHostEnvironment = appContext._wacHostEnvironment || "0";
                returnedContext._isFromWacAutomation = !!appContext._isFromWacAutomation;
                if(OSF.AppTelemetry)
                    OSF.AppTelemetry.initialize(returnedContext);
                gotAppContext(returnedContext)
            }
            else
            {
                var errorMsg = "Function ContextActivationManager_getAppContextAsync call failed. ErrorCode is " + errorCode + ", exception: " + appContext;
                if(OSF.AppTelemetry)
                    OSF.AppTelemetry.logAppException(errorMsg);
                throw errorMsg;
            }
        };
    try
    {
        if(this._hostInfo.isDialog && window.opener != null)
        {
            var appContext = OfficeExt.WACUtils.parseAppContextFromWindowName(false,OSF._OfficeAppFactory.getWindowName());
            getInvocationCallbackWebApp(0,appContext)
        }
        else
            this._webAppState.clientEndPoint.invoke("ContextActivationManager_getAppContextAsync",getInvocationCallbackWebApp,this._webAppState.id)
    }
    catch(ex)
    {
        if(OSF.AppTelemetry)
            OSF.AppTelemetry.logAppException("Exception thrown when trying to invoke getAppContextAsync. Exception:[" + ex + "]");
        throw ex;
    }
};
OSF.InitializationHelper.prototype.setAgaveHostCommunication = function OSF_InitializationHelper$setAgaveHostCommunication()
{
    try
    {
        var me = this;
        var xdmInfoValue = OSF.OUtil.parseXdmInfoWithGivenFragment(false,OSF._OfficeAppFactory.getWindowLocationHash());
        if(!xdmInfoValue && OSF._OfficeAppFactory.getWindowName)
            xdmInfoValue = OSF.OUtil.parseXdmInfoFromWindowName(false,OSF._OfficeAppFactory.getWindowName());
        if(xdmInfoValue)
        {
            var xdmItems = OSF.OUtil.getInfoItems(xdmInfoValue);
            if(xdmItems != undefined && xdmItems.length >= 3)
            {
                me._webAppState.conversationID = xdmItems[0];
                me._webAppState.id = xdmItems[1];
                me._webAppState.webAppUrl = xdmItems[2].indexOf(":") >= 0 ? xdmItems[2] : decodeURIComponent(xdmItems[2])
            }
        }
        me._webAppState.wnd = window.opener != null ? window.opener : window.parent;
        var serializerVersion = OSF.OUtil.parseSerializerVersionWithGivenFragment(false,OSF._OfficeAppFactory.getWindowLocationHash());
        if(isNaN(serializerVersion) && OSF._OfficeAppFactory.getWindowName)
            serializerVersion = OSF.OUtil.parseSerializerVersionFromWindowName(false,OSF._OfficeAppFactory.getWindowName());
        me._webAppState.serializerVersion = serializerVersion;
        if(this._hostInfo.isDialog && window.opener != null)
            return;
        me._webAppState.clientEndPoint = Microsoft.Office.Common.XdmCommunicationManager.connect(me._webAppState.conversationID,me._webAppState.wnd,me._webAppState.webAppUrl,me._webAppState.serializerVersion);
        me._webAppState.serviceEndPoint = Microsoft.Office.Common.XdmCommunicationManager.createServiceEndPoint(me._webAppState.id);
        var notificationConversationId = me._webAppState.conversationID + OSF.SharedConstants.NotificationConversationIdSuffix;
        me._webAppState.serviceEndPoint.registerConversation(notificationConversationId,me._webAppState.webAppUrl);
        var notifyAgave = function OSF__OfficeAppFactory_initialize$notifyAgave(actionId)
            {
                switch(actionId)
                {
                    case OSF.AgaveHostAction.Select:
                        me._webAppState.focused = true;
                        break;
                    case OSF.AgaveHostAction.UnSelect:
                        me._webAppState.focused = false;
                        break;
                    case OSF.AgaveHostAction.TabIn:
                    case OSF.AgaveHostAction.CtrlF6In:
                        window.focus();
                        var list = document.querySelectorAll(me._tabbableElements);
                        var focused = OSF.OUtil.focusToFirstTabbable(list,false);
                        if(!focused)
                        {
                            window.blur();
                            me._webAppState.focused = false;
                            me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost",null,[me._webAppState.id,OSF.AgaveHostAction.ExitNoFocusable])
                        }
                        break;
                    case OSF.AgaveHostAction.TabInShift:
                        window.focus();
                        var list = document.querySelectorAll(me._tabbableElements);
                        var focused = OSF.OUtil.focusToFirstTabbable(list,true);
                        if(!focused)
                        {
                            window.blur();
                            me._webAppState.focused = false;
                            me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost",null,[me._webAppState.id,OSF.AgaveHostAction.ExitNoFocusableShift])
                        }
                        break;
                    default:
                        OsfMsAjaxFactory.msAjaxDebug.trace("actionId " + actionId + " notifyAgave is wrong.");
                        break
                }
            };
        me._webAppState.serviceEndPoint.registerMethod("Office_notifyAgave",notifyAgave,Microsoft.Office.Common.InvokeType.async,false);
        me.addOrRemoveEventListenersForWindow(true)
    }
    catch(ex)
    {
        if(OSF.AppTelemetry)
            OSF.AppTelemetry.logAppException("Exception thrown in setAgaveHostCommunication. Exception:[" + ex + "]");
        throw ex;
    }
};
OSF.InitializationHelper.prototype.addOrRemoveEventListenersForWindow = function OSF_InitializationHelper$addOrRemoveEventListenersForWindow(isAdd)
{
    var me = this;
    var onWindowFocus = function()
        {
            if(!me._webAppState.focused)
                me._webAppState.focused = true;
            me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost",null,[me._webAppState.id,OSF.AgaveHostAction.Select])
        };
    var onWindowBlur = function()
        {
            if(!OSF)
                return;
            if(me._webAppState.focused)
                me._webAppState.focused = false;
            me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost",null,[me._webAppState.id,OSF.AgaveHostAction.UnSelect])
        };
    var onWindowKeydown = function(e)
        {
            e.preventDefault = e.preventDefault || function()
            {
                e.returnValue = false
            };
            if(e.keyCode == 117 && (e.ctrlKey || e.metaKey))
            {
                var actionId = OSF.AgaveHostAction.CtrlF6Exit;
                if(e.shiftKey)
                    actionId = OSF.AgaveHostAction.CtrlF6ExitShift;
                me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost",null,[me._webAppState.id,actionId])
            }
            else if(e.keyCode == 9)
            {
                e.preventDefault();
                var allTabbableElements = document.querySelectorAll(me._tabbableElements);
                var focused = OSF.OUtil.focusToNextTabbable(allTabbableElements,e.target || e.srcElement,e.shiftKey);
                if(!focused)
                    if(me._hostInfo.isDialog)
                        OSF.OUtil.focusToFirstTabbable(allTabbableElements,e.shiftKey);
                    else if(e.shiftKey)
                        me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost",null,[me._webAppState.id,OSF.AgaveHostAction.TabExitShift]);
                    else
                        me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost",null,[me._webAppState.id,OSF.AgaveHostAction.TabExit])
            }
            else if(e.keyCode == 27)
            {
                e.preventDefault();
                me.dismissDialogNotification && me.dismissDialogNotification();
                me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost",null,[me._webAppState.id,OSF.AgaveHostAction.EscExit])
            }
            else if(e.keyCode == 113)
            {
                e.preventDefault();
                me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost",null,[me._webAppState.id,OSF.AgaveHostAction.F2Exit])
            }
        };
    var onWindowKeypress = function(e)
        {
            if(e.keyCode == 117 && e.ctrlKey)
                if(e.preventDefault)
                    e.preventDefault();
                else
                    e.returnValue = false
        };
    if(isAdd)
    {
        OSF.OUtil.addEventListener(window,"focus",onWindowFocus);
        OSF.OUtil.addEventListener(window,"blur",onWindowBlur);
        OSF.OUtil.addEventListener(window,"keydown",onWindowKeydown);
        OSF.OUtil.addEventListener(window,"keypress",onWindowKeypress)
    }
    else
    {
        OSF.OUtil.removeEventListener(window,"focus",onWindowFocus);
        OSF.OUtil.removeEventListener(window,"blur",onWindowBlur);
        OSF.OUtil.removeEventListener(window,"keydown",onWindowKeydown);
        OSF.OUtil.removeEventListener(window,"keypress",onWindowKeypress)
    }
};
OSF.InitializationHelper.prototype.initWebDialog = function OSF_InitializationHelper$initWebDialog(appContext)
{
    if(appContext.get_isDialog())
    {
        if(OSF.DDA.UI.ChildUI)
        {
            var isPopupWindow = window.opener != null;
            appContext.ui = new OSF.DDA.UI.ChildUI(isPopupWindow);
            if(isPopupWindow)
                this.registerMessageReceivedEventForWindowDialog && this.registerMessageReceivedEventForWindowDialog()
        }
    }
    else if(OSF.DDA.UI.ParentUI)
    {
        appContext.ui = new OSF.DDA.UI.ParentUI;
        if(OfficeExt.Container)
            OSF.DDA.DispIdHost.addAsyncMethods(appContext.ui,[OSF.DDA.AsyncMethodNames.CloseContainerAsync])
    }
};
OSF.InitializationHelper.prototype.initWebAuth = function OSF_InitializationHelper$initWebAuth(appContext)
{
    if(OSF.DDA.Auth)
    {
        appContext.auth = new OSF.DDA.Auth;
        OSF.DDA.DispIdHost.addAsyncMethods(appContext.auth,[OSF.DDA.AsyncMethodNames.GetAccessTokenAsync])
    }
};
OSF.InitializationHelper.prototype.initWebAuthImplicit = function OSF_InitializationHelper$initWebAuthImplicit(appContext)
{
    if(OSF.DDA.WebAuth)
    {
        appContext.webAuth = new OSF.DDA.WebAuth;
        OSF.DDA.DispIdHost.addAsyncMethods(appContext.webAuth,[OSF.DDA.AsyncMethodNames.GetAuthContextAsync])
    }
};
OSF.getClientEndPoint = function OSF$getClientEndPoint()
{
    var initializationHelper = OSF._OfficeAppFactory.getInitializationHelper();
    return initializationHelper._webAppState.clientEndPoint
};
OSF.InitializationHelper.prototype.prepareRightAfterWebExtensionInitialize = function OSF_InitializationHelper$prepareRightAfterWebExtensionInitialize()
{
    if(this._hostInfo.isDialog)
    {
        window.focus();
        var list = document.querySelectorAll(this._tabbableElements);
        var focused = OSF.OUtil.focusToFirstTabbable(list,false);
        if(!focused)
        {
            window.blur();
            this._webAppState.focused = false;
            if(this._webAppState.clientEndPoint)
                this._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost",null,[this._webAppState.id,OSF.AgaveHostAction.ExitNoFocusable])
        }
    }
};
(function()
{
    var checkScriptOverride = function OSF$checkScriptOverride()
        {
            var postScriptOverrideCheckAction = function OSF$postScriptOverrideCheckAction(customizedScriptPath)
                {
                    if(customizedScriptPath)
                        OSF.OUtil.loadScript(customizedScriptPath,function()
                        {
                            OsfMsAjaxFactory.msAjaxDebug.trace("loaded customized script:" + customizedScriptPath)
                        })
                };
            var conversationID,
                webAppUrl,
                items;
            var clientEndPoint = null;
            var xdmInfoValue = OSF.OUtil.parseXdmInfo();
            if(xdmInfoValue)
            {
                items = OSF.OUtil.getInfoItems(xdmInfoValue);
                if(items && items.length >= 3)
                {
                    conversationID = items[0];
                    webAppUrl = items[2];
                    var serializerVersion = OSF.OUtil.parseSerializerVersionWithGivenFragment(false,OSF._OfficeAppFactory.getWindowLocationHash());
                    if(isNaN(serializerVersion) && OSF._OfficeAppFactory.getWindowName)
                        serializerVersion = OSF.OUtil.parseSerializerVersionFromWindowName(false,OSF._OfficeAppFactory.getWindowName());
                    clientEndPoint = Microsoft.Office.Common.XdmCommunicationManager.connect(conversationID,window.parent,webAppUrl,serializerVersion)
                }
            }
            var customizedScriptPath = null;
            if(!clientEndPoint)
            {
                try
                {
                    if(window.external && typeof window.external.getCustomizedScriptPath !== "undefined")
                        customizedScriptPath = window.external.getCustomizedScriptPath()
                }
                catch(ex)
                {
                    OsfMsAjaxFactory.msAjaxDebug.trace("no script override through window.external.")
                }
                postScriptOverrideCheckAction(customizedScriptPath)
            }
            else
                try
                {
                    clientEndPoint.invoke("getCustomizedScriptPathAsync",function OSF$getCustomizedScriptPathAsyncCallback(errorCode, scriptPath)
                    {
                        postScriptOverrideCheckAction(errorCode === 0 ? scriptPath : null)
                    },{__timeout__: 1e3})
                }
                catch(ex)
                {
                    OsfMsAjaxFactory.msAjaxDebug.trace("no script override through cross frame communication.")
                }
        };
    var requiresMsAjax = true;
    if(requiresMsAjax && !OsfMsAjaxFactory.isMsAjaxLoaded())
        if(!(OSF._OfficeAppFactory && OSF._OfficeAppFactory && OSF._OfficeAppFactory.getLoadScriptHelper && OSF._OfficeAppFactory.getLoadScriptHelper().isScriptLoading(OSF.ConstantNames.MicrosoftAjaxId)))
            OsfMsAjaxFactory.loadMsAjaxFull(function OSF$loadMSAjaxCallback()
            {
                if(OsfMsAjaxFactory.isMsAjaxLoaded())
                    checkScriptOverride();
                else
                    throw"Not able to load MicrosoftAjax.js.";
            });
        else
            OSF._OfficeAppFactory.getLoadScriptHelper().waitForScripts([OSF.ConstantNames.MicrosoftAjaxId],checkScriptOverride);
    else
        checkScriptOverride()
})();
var OSFLog;
(function(OSFLog)
{
    var BaseUsageData = function()
        {
            function BaseUsageData(table)
            {
                this._table = table;
                this._fields = {}
            }
            Object.defineProperty(BaseUsageData.prototype,"Fields",{
                get: function()
                {
                    return this._fields
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(BaseUsageData.prototype,"Table",{
                get: function()
                {
                    return this._table
                },
                enumerable: true,
                configurable: true
            });
            BaseUsageData.prototype.SerializeFields = function(){};
            BaseUsageData.prototype.SetSerializedField = function(key, value)
            {
                if(typeof value !== "undefined" && value !== null)
                    this._serializedFields[key] = value.toString()
            };
            BaseUsageData.prototype.SerializeRow = function()
            {
                this._serializedFields = {};
                this.SetSerializedField("Table",this._table);
                this.SerializeFields();
                return JSON.stringify(this._serializedFields)
            };
            return BaseUsageData
        }();
    OSFLog.BaseUsageData = BaseUsageData;
    var AppActivatedUsageData = function(_super)
        {
            __extends(AppActivatedUsageData,_super);
            function AppActivatedUsageData()
            {
                _super.call(this,"AppActivated")
            }
            Object.defineProperty(AppActivatedUsageData.prototype,"CorrelationId",{
                get: function()
                {
                    return this.Fields["CorrelationId"]
                },
                set: function(value)
                {
                    this.Fields["CorrelationId"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppActivatedUsageData.prototype,"SessionId",{
                get: function()
                {
                    return this.Fields["SessionId"]
                },
                set: function(value)
                {
                    this.Fields["SessionId"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppActivatedUsageData.prototype,"AppId",{
                get: function()
                {
                    return this.Fields["AppId"]
                },
                set: function(value)
                {
                    this.Fields["AppId"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppActivatedUsageData.prototype,"AppInstanceId",{
                get: function()
                {
                    return this.Fields["AppInstanceId"]
                },
                set: function(value)
                {
                    this.Fields["AppInstanceId"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppActivatedUsageData.prototype,"AppURL",{
                get: function()
                {
                    return this.Fields["AppURL"]
                },
                set: function(value)
                {
                    this.Fields["AppURL"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppActivatedUsageData.prototype,"AssetId",{
                get: function()
                {
                    return this.Fields["AssetId"]
                },
                set: function(value)
                {
                    this.Fields["AssetId"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppActivatedUsageData.prototype,"Browser",{
                get: function()
                {
                    return this.Fields["Browser"]
                },
                set: function(value)
                {
                    this.Fields["Browser"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppActivatedUsageData.prototype,"UserId",{
                get: function()
                {
                    return this.Fields["UserId"]
                },
                set: function(value)
                {
                    this.Fields["UserId"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppActivatedUsageData.prototype,"Host",{
                get: function()
                {
                    return this.Fields["Host"]
                },
                set: function(value)
                {
                    this.Fields["Host"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppActivatedUsageData.prototype,"HostVersion",{
                get: function()
                {
                    return this.Fields["HostVersion"]
                },
                set: function(value)
                {
                    this.Fields["HostVersion"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppActivatedUsageData.prototype,"ClientId",{
                get: function()
                {
                    return this.Fields["ClientId"]
                },
                set: function(value)
                {
                    this.Fields["ClientId"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppActivatedUsageData.prototype,"AppSizeWidth",{
                get: function()
                {
                    return this.Fields["AppSizeWidth"]
                },
                set: function(value)
                {
                    this.Fields["AppSizeWidth"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppActivatedUsageData.prototype,"AppSizeHeight",{
                get: function()
                {
                    return this.Fields["AppSizeHeight"]
                },
                set: function(value)
                {
                    this.Fields["AppSizeHeight"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppActivatedUsageData.prototype,"Message",{
                get: function()
                {
                    return this.Fields["Message"]
                },
                set: function(value)
                {
                    this.Fields["Message"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppActivatedUsageData.prototype,"DocUrl",{
                get: function()
                {
                    return this.Fields["DocUrl"]
                },
                set: function(value)
                {
                    this.Fields["DocUrl"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppActivatedUsageData.prototype,"OfficeJSVersion",{
                get: function()
                {
                    return this.Fields["OfficeJSVersion"]
                },
                set: function(value)
                {
                    this.Fields["OfficeJSVersion"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppActivatedUsageData.prototype,"HostJSVersion",{
                get: function()
                {
                    return this.Fields["HostJSVersion"]
                },
                set: function(value)
                {
                    this.Fields["HostJSVersion"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppActivatedUsageData.prototype,"WacHostEnvironment",{
                get: function()
                {
                    return this.Fields["WacHostEnvironment"]
                },
                set: function(value)
                {
                    this.Fields["WacHostEnvironment"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppActivatedUsageData.prototype,"IsFromWacAutomation",{
                get: function()
                {
                    return this.Fields["IsFromWacAutomation"]
                },
                set: function(value)
                {
                    this.Fields["IsFromWacAutomation"] = value
                },
                enumerable: true,
                configurable: true
            });
            AppActivatedUsageData.prototype.SerializeFields = function()
            {
                this.SetSerializedField("CorrelationId",this.CorrelationId);
                this.SetSerializedField("SessionId",this.SessionId);
                this.SetSerializedField("AppId",this.AppId);
                this.SetSerializedField("AppInstanceId",this.AppInstanceId);
                this.SetSerializedField("AppURL",this.AppURL);
                this.SetSerializedField("AssetId",this.AssetId);
                this.SetSerializedField("Browser",this.Browser);
                this.SetSerializedField("UserId",this.UserId);
                this.SetSerializedField("Host",this.Host);
                this.SetSerializedField("HostVersion",this.HostVersion);
                this.SetSerializedField("ClientId",this.ClientId);
                this.SetSerializedField("AppSizeWidth",this.AppSizeWidth);
                this.SetSerializedField("AppSizeHeight",this.AppSizeHeight);
                this.SetSerializedField("Message",this.Message);
                this.SetSerializedField("DocUrl",this.DocUrl);
                this.SetSerializedField("OfficeJSVersion",this.OfficeJSVersion);
                this.SetSerializedField("HostJSVersion",this.HostJSVersion);
                this.SetSerializedField("WacHostEnvironment",this.WacHostEnvironment);
                this.SetSerializedField("IsFromWacAutomation",this.IsFromWacAutomation)
            };
            return AppActivatedUsageData
        }(BaseUsageData);
    OSFLog.AppActivatedUsageData = AppActivatedUsageData;
    var ScriptLoadUsageData = function(_super)
        {
            __extends(ScriptLoadUsageData,_super);
            function ScriptLoadUsageData()
            {
                _super.call(this,"ScriptLoad")
            }
            Object.defineProperty(ScriptLoadUsageData.prototype,"CorrelationId",{
                get: function()
                {
                    return this.Fields["CorrelationId"]
                },
                set: function(value)
                {
                    this.Fields["CorrelationId"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(ScriptLoadUsageData.prototype,"SessionId",{
                get: function()
                {
                    return this.Fields["SessionId"]
                },
                set: function(value)
                {
                    this.Fields["SessionId"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(ScriptLoadUsageData.prototype,"ScriptId",{
                get: function()
                {
                    return this.Fields["ScriptId"]
                },
                set: function(value)
                {
                    this.Fields["ScriptId"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(ScriptLoadUsageData.prototype,"StartTime",{
                get: function()
                {
                    return this.Fields["StartTime"]
                },
                set: function(value)
                {
                    this.Fields["StartTime"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(ScriptLoadUsageData.prototype,"ResponseTime",{
                get: function()
                {
                    return this.Fields["ResponseTime"]
                },
                set: function(value)
                {
                    this.Fields["ResponseTime"] = value
                },
                enumerable: true,
                configurable: true
            });
            ScriptLoadUsageData.prototype.SerializeFields = function()
            {
                this.SetSerializedField("CorrelationId",this.CorrelationId);
                this.SetSerializedField("SessionId",this.SessionId);
                this.SetSerializedField("ScriptId",this.ScriptId);
                this.SetSerializedField("StartTime",this.StartTime);
                this.SetSerializedField("ResponseTime",this.ResponseTime)
            };
            return ScriptLoadUsageData
        }(BaseUsageData);
    OSFLog.ScriptLoadUsageData = ScriptLoadUsageData;
    var AppClosedUsageData = function(_super)
        {
            __extends(AppClosedUsageData,_super);
            function AppClosedUsageData()
            {
                _super.call(this,"AppClosed")
            }
            Object.defineProperty(AppClosedUsageData.prototype,"CorrelationId",{
                get: function()
                {
                    return this.Fields["CorrelationId"]
                },
                set: function(value)
                {
                    this.Fields["CorrelationId"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppClosedUsageData.prototype,"SessionId",{
                get: function()
                {
                    return this.Fields["SessionId"]
                },
                set: function(value)
                {
                    this.Fields["SessionId"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppClosedUsageData.prototype,"FocusTime",{
                get: function()
                {
                    return this.Fields["FocusTime"]
                },
                set: function(value)
                {
                    this.Fields["FocusTime"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppClosedUsageData.prototype,"AppSizeFinalWidth",{
                get: function()
                {
                    return this.Fields["AppSizeFinalWidth"]
                },
                set: function(value)
                {
                    this.Fields["AppSizeFinalWidth"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppClosedUsageData.prototype,"AppSizeFinalHeight",{
                get: function()
                {
                    return this.Fields["AppSizeFinalHeight"]
                },
                set: function(value)
                {
                    this.Fields["AppSizeFinalHeight"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppClosedUsageData.prototype,"OpenTime",{
                get: function()
                {
                    return this.Fields["OpenTime"]
                },
                set: function(value)
                {
                    this.Fields["OpenTime"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppClosedUsageData.prototype,"CloseMethod",{
                get: function()
                {
                    return this.Fields["CloseMethod"]
                },
                set: function(value)
                {
                    this.Fields["CloseMethod"] = value
                },
                enumerable: true,
                configurable: true
            });
            AppClosedUsageData.prototype.SerializeFields = function()
            {
                this.SetSerializedField("CorrelationId",this.CorrelationId);
                this.SetSerializedField("SessionId",this.SessionId);
                this.SetSerializedField("FocusTime",this.FocusTime);
                this.SetSerializedField("AppSizeFinalWidth",this.AppSizeFinalWidth);
                this.SetSerializedField("AppSizeFinalHeight",this.AppSizeFinalHeight);
                this.SetSerializedField("OpenTime",this.OpenTime);
                this.SetSerializedField("CloseMethod",this.CloseMethod)
            };
            return AppClosedUsageData
        }(BaseUsageData);
    OSFLog.AppClosedUsageData = AppClosedUsageData;
    var APIUsageUsageData = function(_super)
        {
            __extends(APIUsageUsageData,_super);
            function APIUsageUsageData()
            {
                _super.call(this,"APIUsage")
            }
            Object.defineProperty(APIUsageUsageData.prototype,"CorrelationId",{
                get: function()
                {
                    return this.Fields["CorrelationId"]
                },
                set: function(value)
                {
                    this.Fields["CorrelationId"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(APIUsageUsageData.prototype,"SessionId",{
                get: function()
                {
                    return this.Fields["SessionId"]
                },
                set: function(value)
                {
                    this.Fields["SessionId"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(APIUsageUsageData.prototype,"APIType",{
                get: function()
                {
                    return this.Fields["APIType"]
                },
                set: function(value)
                {
                    this.Fields["APIType"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(APIUsageUsageData.prototype,"APIID",{
                get: function()
                {
                    return this.Fields["APIID"]
                },
                set: function(value)
                {
                    this.Fields["APIID"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(APIUsageUsageData.prototype,"Parameters",{
                get: function()
                {
                    return this.Fields["Parameters"]
                },
                set: function(value)
                {
                    this.Fields["Parameters"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(APIUsageUsageData.prototype,"ResponseTime",{
                get: function()
                {
                    return this.Fields["ResponseTime"]
                },
                set: function(value)
                {
                    this.Fields["ResponseTime"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(APIUsageUsageData.prototype,"ErrorType",{
                get: function()
                {
                    return this.Fields["ErrorType"]
                },
                set: function(value)
                {
                    this.Fields["ErrorType"] = value
                },
                enumerable: true,
                configurable: true
            });
            APIUsageUsageData.prototype.SerializeFields = function()
            {
                this.SetSerializedField("CorrelationId",this.CorrelationId);
                this.SetSerializedField("SessionId",this.SessionId);
                this.SetSerializedField("APIType",this.APIType);
                this.SetSerializedField("APIID",this.APIID);
                this.SetSerializedField("Parameters",this.Parameters);
                this.SetSerializedField("ResponseTime",this.ResponseTime);
                this.SetSerializedField("ErrorType",this.ErrorType)
            };
            return APIUsageUsageData
        }(BaseUsageData);
    OSFLog.APIUsageUsageData = APIUsageUsageData;
    var AppInitializationUsageData = function(_super)
        {
            __extends(AppInitializationUsageData,_super);
            function AppInitializationUsageData()
            {
                _super.call(this,"AppInitialization")
            }
            Object.defineProperty(AppInitializationUsageData.prototype,"CorrelationId",{
                get: function()
                {
                    return this.Fields["CorrelationId"]
                },
                set: function(value)
                {
                    this.Fields["CorrelationId"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppInitializationUsageData.prototype,"SessionId",{
                get: function()
                {
                    return this.Fields["SessionId"]
                },
                set: function(value)
                {
                    this.Fields["SessionId"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppInitializationUsageData.prototype,"SuccessCode",{
                get: function()
                {
                    return this.Fields["SuccessCode"]
                },
                set: function(value)
                {
                    this.Fields["SuccessCode"] = value
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(AppInitializationUsageData.prototype,"Message",{
                get: function()
                {
                    return this.Fields["Message"]
                },
                set: function(value)
                {
                    this.Fields["Message"] = value
                },
                enumerable: true,
                configurable: true
            });
            AppInitializationUsageData.prototype.SerializeFields = function()
            {
                this.SetSerializedField("CorrelationId",this.CorrelationId);
                this.SetSerializedField("SessionId",this.SessionId);
                this.SetSerializedField("SuccessCode",this.SuccessCode);
                this.SetSerializedField("Message",this.Message)
            };
            return AppInitializationUsageData
        }(BaseUsageData);
    OSFLog.AppInitializationUsageData = AppInitializationUsageData
})(OSFLog || (OSFLog = {}));
var Logger;
(function(Logger)
{
    "use strict";
    (function(TraceLevel)
    {
        TraceLevel[TraceLevel["info"] = 0] = "info";
        TraceLevel[TraceLevel["warning"] = 1] = "warning";
        TraceLevel[TraceLevel["error"] = 2] = "error"
    })(Logger.TraceLevel || (Logger.TraceLevel = {}));
    var TraceLevel = Logger.TraceLevel;
    (function(SendFlag)
    {
        SendFlag[SendFlag["none"] = 0] = "none";
        SendFlag[SendFlag["flush"] = 1] = "flush"
    })(Logger.SendFlag || (Logger.SendFlag = {}));
    var SendFlag = Logger.SendFlag;
    function allowUploadingData(){}
    Logger.allowUploadingData = allowUploadingData;
    function sendLog(traceLevel, message, flag){}
    Logger.sendLog = sendLog;
    function creatULSEndpoint()
    {
        try
        {
            return new ULSEndpointProxy
        }
        catch(e)
        {
            return null
        }
    }
    var ULSEndpointProxy = function()
        {
            function ULSEndpointProxy(){}
            ULSEndpointProxy.prototype.writeLog = function(log){};
            ULSEndpointProxy.prototype.loadProxyFrame = function(){};
            return ULSEndpointProxy
        }();
    if(!OSF.Logger)
        OSF.Logger = Logger;
    Logger.ulsEndpoint = creatULSEndpoint()
})(Logger || (Logger = {}));
var OSFAriaLogger;
(function(OSFAriaLogger)
{
    var TelemetryEventAppActivated = {
            name: "AppActivated",
            enabled: true,
            basic: true,
            critical: true,
            points: [{
                    name: "Browser",
                    type: "string"
                },{
                    name: "Message",
                    type: "string"
                },{
                    name: "AppURL",
                    type: "string"
                },{
                    name: "Host",
                    type: "string"
                },{
                    name: "AppSizeWidth",
                    type: "int64"
                },{
                    name: "AppSizeHeight",
                    type: "int64"
                },{
                    name: "IsFromWacAutomation",
                    type: "string"
                },]
        };
    var TelemetryEventScriptLoad = {
            name: "ScriptLoad",
            enabled: true,
            basic: false,
            critical: false,
            points: [{
                    name: "ScriptId",
                    type: "string"
                },{
                    name: "StartTime",
                    type: "double"
                },{
                    name: "ResponseTime",
                    type: "double"
                },]
        };
    var TelemetryEventApiUsage = {
            name: "APIUsage",
            enabled: false,
            basic: false,
            critical: false,
            points: [{
                    name: "APIType",
                    type: "string"
                },{
                    name: "APIID",
                    type: "int64"
                },{
                    name: "Parameters",
                    type: "string"
                },{
                    name: "ResponseTime",
                    type: "int64"
                },{
                    name: "ErrorType",
                    type: "int64"
                },]
        };
    var TelemetryEventAppInitialization = {
            name: "AppInitialization",
            enabled: true,
            basic: false,
            critical: false,
            points: [{
                    name: "SuccessCode",
                    type: "int64"
                },{
                    name: "Message",
                    type: "string"
                },]
        };
    var TelemetryEventAppClosed = {
            name: "AppClosed",
            enabled: true,
            basic: false,
            critical: false,
            points: [{
                    name: "FocusTime",
                    type: "int64"
                },{
                    name: "AppSizeFinalWidth",
                    type: "int64"
                },{
                    name: "AppSizeFinalHeight",
                    type: "int64"
                },{
                    name: "OpenTime",
                    type: "int64"
                },]
        };
    var TelemetryEvents = [TelemetryEventAppActivated,TelemetryEventScriptLoad,TelemetryEventApiUsage,TelemetryEventAppInitialization,TelemetryEventAppClosed,];
    function createDataField(value, point)
    {
        var key = point.rename === undefined ? point.name : point.rename;
        var type = point.type;
        var field = undefined;
        switch(type)
        {
            case"string":
                field = oteljs.makeStringDataField(key,value);
                break;
            case"double":
                if(typeof value === "string")
                    value = parseFloat(value);
                field = oteljs.makeDoubleDataField(key,value);
                break;
            case"int64":
                if(typeof value === "string")
                    value = parseInt(value);
                field = oteljs.makeInt64DataField(key,value);
                break;
            case"boolean":
                if(typeof value === "string")
                    value = value === "true";
                field = oteljs.makeBooleanDataField(key,value);
                break
        }
        return field
    }
    function getEventDefinition(eventName)
    {
        for(var _i = 0; _i < TelemetryEvents.length; _i++)
        {
            var event_1 = TelemetryEvents[_i];
            if(event_1.name === eventName)
                return event_1
        }
        return undefined
    }
    function eventEnabled(eventName)
    {
        var eventDefinition = getEventDefinition(eventName);
        if(eventDefinition === undefined)
            return false;
        return eventDefinition.enabled
    }
    function generateTelemetryEvent(eventName, telemetryData)
    {
        var eventDefinition = getEventDefinition(eventName);
        if(eventDefinition === undefined)
            return undefined;
        var dataFields = [];
        for(var _i = 0, _a = eventDefinition.points; _i < _a.length; _i++)
        {
            var point = _a[_i];
            var key = point.name;
            var value = telemetryData[key];
            if(value === undefined)
                continue;
            var field = createDataField(value,point);
            if(field !== undefined)
                dataFields.push(field)
        }
        var flags = {dataCategories: oteljs.DataCategories.ProductServiceUsage};
        if(eventDefinition.critical)
            flags.samplingPolicy = oteljs.SamplingPolicy.CriticalBusinessImpact;
        if(eventDefinition.basic)
            flags.diagnosticLevel = oteljs.DiagnosticLevel.BasicEvent;
        var eventNameFull = "Office.Extensibility.OfficeJs." + eventName + "X";
        var event = {
                eventName: eventNameFull,
                dataFields: dataFields,
                eventFlags: flags
            };
        return event
    }
    function sendOtelTelemetryEvent(eventName, telemetryData)
    {
        if(eventEnabled(eventName))
            if(typeof OTel !== "undefined")
                OTel.OTelLogger.onTelemetryLoaded(function()
                {
                    var event = generateTelemetryEvent(eventName,telemetryData);
                    if(event === undefined)
                        return;
                    Microsoft.Office.WebExtension.sendTelemetryEvent(event)
                })
    }
    var AriaLogger = function()
        {
            function AriaLogger(){}
            AriaLogger.prototype.getAriaCDNLocation = function()
            {
                return OSF._OfficeAppFactory.getLoadScriptHelper().getOfficeJsBasePath() + "ariatelemetry/aria-web-telemetry.js"
            };
            AriaLogger.getInstance = function()
            {
                if(AriaLogger.AriaLoggerObj === undefined)
                    AriaLogger.AriaLoggerObj = new AriaLogger;
                return AriaLogger.AriaLoggerObj
            };
            AriaLogger.prototype.isIUsageData = function(arg)
            {
                return arg["Fields"] !== undefined
            };
            AriaLogger.prototype.sendTelemetry = function(tableName, telemetryData)
            {
                var startAfterMs = 1e3;
                if(AriaLogger.EnableSendingTelemetryWithLegacyAria)
                    OSF.OUtil.loadScript(this.getAriaCDNLocation(),function()
                    {
                        try
                        {
                            if(!this.ALogger)
                            {
                                var OfficeExtensibilityTenantID = "db334b301e7b474db5e0f02f07c51a47-a1b5bc36-1bbe-482f-a64a-c2d9cb606706-7439";
                                this.ALogger = AWTLogManager.initialize(OfficeExtensibilityTenantID)
                            }
                            var eventProperties = new AWTEventProperties;
                            eventProperties.setName("Office.Extensibility.OfficeJS." + tableName);
                            for(var key in telemetryData)
                                if(key.toLowerCase() !== "table")
                                    eventProperties.setProperty(key,telemetryData[key]);
                            var today = new Date;
                            eventProperties.setProperty("Date",today.toISOString());
                            this.ALogger.logEvent(eventProperties)
                        }
                        catch(e){}
                    },startAfterMs);
                if(AriaLogger.EnableSendingTelemetryWithOTel)
                    sendOtelTelemetryEvent(tableName,telemetryData)
            };
            AriaLogger.prototype.logData = function(data)
            {
                if(this.isIUsageData(data))
                    this.sendTelemetry(data["Table"],data["Fields"]);
                else
                    this.sendTelemetry(data["Table"],data)
            };
            AriaLogger.EnableSendingTelemetryWithOTel = true;
            AriaLogger.EnableSendingTelemetryWithLegacyAria = true;
            return AriaLogger
        }();
    OSFAriaLogger.AriaLogger = AriaLogger
})(OSFAriaLogger || (OSFAriaLogger = {}));
var OSFAppTelemetry;
(function(OSFAppTelemetry)
{
    "use strict";
    var appInfo;
    var sessionId = OSF.OUtil.Guid.generateNewGuid();
    var osfControlAppCorrelationId = "";
    var omexDomainRegex = new RegExp("^https?://store\\.office(ppe|-int)?\\.com/","i");
    OSFAppTelemetry.enableTelemetry = true;
    var AppInfo = function()
        {
            function AppInfo(){}
            return AppInfo
        }();
    OSFAppTelemetry.AppInfo = AppInfo;
    var Event = function()
        {
            function Event(name, handler)
            {
                this.name = name;
                this.handler = handler
            }
            return Event
        }();
    var AppStorage = function()
        {
            function AppStorage()
            {
                this.clientIDKey = "Office API client";
                this.logIdSetKey = "Office App Log Id Set"
            }
            AppStorage.prototype.getClientId = function()
            {
                var clientId = this.getValue(this.clientIDKey);
                if(!clientId || clientId.length <= 0 || clientId.length > 40)
                {
                    clientId = OSF.OUtil.Guid.generateNewGuid();
                    this.setValue(this.clientIDKey,clientId)
                }
                return clientId
            };
            AppStorage.prototype.saveLog = function(logId, log)
            {
                var logIdSet = this.getValue(this.logIdSetKey);
                logIdSet = (logIdSet && logIdSet.length > 0 ? logIdSet + ";" : "") + logId;
                this.setValue(this.logIdSetKey,logIdSet);
                this.setValue(logId,log)
            };
            AppStorage.prototype.enumerateLog = function(callback, clean)
            {
                var logIdSet = this.getValue(this.logIdSetKey);
                if(logIdSet)
                {
                    var ids = logIdSet.split(";");
                    for(var id in ids)
                    {
                        var logId = ids[id];
                        var log = this.getValue(logId);
                        if(log)
                        {
                            if(callback)
                                callback(logId,log);
                            if(clean)
                                this.remove(logId)
                        }
                    }
                    if(clean)
                        this.remove(this.logIdSetKey)
                }
            };
            AppStorage.prototype.getValue = function(key)
            {
                var osfLocalStorage = OSF.OUtil.getLocalStorage();
                var value = "";
                if(osfLocalStorage)
                    value = osfLocalStorage.getItem(key);
                return value
            };
            AppStorage.prototype.setValue = function(key, value)
            {
                var osfLocalStorage = OSF.OUtil.getLocalStorage();
                if(osfLocalStorage)
                    osfLocalStorage.setItem(key,value)
            };
            AppStorage.prototype.remove = function(key)
            {
                var osfLocalStorage = OSF.OUtil.getLocalStorage();
                if(osfLocalStorage)
                    try
                    {
                        osfLocalStorage.removeItem(key)
                    }
                    catch(ex){}
            };
            return AppStorage
        }();
    var AppLogger = function()
        {
            function AppLogger(){}
            AppLogger.prototype.LogData = function(data)
            {
                if(!OSFAppTelemetry.enableTelemetry)
                    return;
                try
                {
                    OSFAriaLogger.AriaLogger.getInstance().logData(data)
                }
                catch(e){}
            };
            AppLogger.prototype.LogRawData = function(log)
            {
                if(!OSFAppTelemetry.enableTelemetry)
                    return;
                try
                {
                    OSFAriaLogger.AriaLogger.getInstance().logData(JSON.parse(log))
                }
                catch(e){}
            };
            return AppLogger
        }();
    function trimStringToLowerCase(input)
    {
        if(input)
            input = input.replace(/[{}]/g,"").toLowerCase();
        return input || ""
    }
    var UrlFilter = function()
        {
            function UrlFilter(){}
            UrlFilter.hashString = function(s)
            {
                var hash = 0;
                if(s.length === 0)
                    return hash;
                for(var i = 0; i < s.length; i++)
                {
                    var c = s.charCodeAt(i);
                    hash = (hash << 5) - hash + c;
                    hash |= 0
                }
                return hash
            };
            UrlFilter.stringToHash = function(s)
            {
                var hash = UrlFilter.hashString(s);
                var stringHash = hash.toString();
                if(hash < 0)
                    stringHash = "1" + stringHash.substring(1);
                else
                    stringHash = "0" + stringHash;
                return stringHash
            };
            UrlFilter.startsWith = function(s, prefix)
            {
                return s.indexOf(prefix) == -0
            };
            UrlFilter.isFileUrl = function(url)
            {
                return UrlFilter.startsWith(url.toLowerCase(),"file:")
            };
            UrlFilter.removeHttpPrefix = function(url)
            {
                var prefix = "";
                if(UrlFilter.startsWith(url.toLowerCase(),UrlFilter.httpsPrefix))
                    prefix = UrlFilter.httpsPrefix;
                else if(UrlFilter.startsWith(url.toLowerCase(),UrlFilter.httpPrefix))
                    prefix = UrlFilter.httpPrefix;
                var clean = url.slice(prefix.length);
                return clean
            };
            UrlFilter.getUrlDomain = function(url)
            {
                var domain = UrlFilter.removeHttpPrefix(url);
                domain = domain.split("/")[0];
                domain = domain.split(":")[0];
                return domain
            };
            UrlFilter.isIp4Address = function(domain)
            {
                var ipv4Regex = /^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$/;
                return ipv4Regex.test(domain)
            };
            UrlFilter.filter = function(url)
            {
                if(UrlFilter.isFileUrl(url))
                {
                    var hash = UrlFilter.stringToHash(url);
                    return"file://" + hash
                }
                var domain = UrlFilter.getUrlDomain(url);
                if(UrlFilter.isIp4Address(domain))
                {
                    var hash = UrlFilter.stringToHash(url);
                    if(UrlFilter.startsWith(domain,"10."))
                        return"IP10Range_" + hash;
                    else if(UrlFilter.startsWith(domain,"192."))
                        return"IP192Range_" + hash;
                    else if(UrlFilter.startsWith(domain,"127."))
                        return"IP127Range_" + hash;
                    return"IPOther_" + hash
                }
                return domain
            };
            UrlFilter.httpPrefix = "http://";
            UrlFilter.httpsPrefix = "https://";
            return UrlFilter
        }();
    function initialize(context)
    {
        if(!OSFAppTelemetry.enableTelemetry)
            return;
        if(appInfo)
            return;
        appInfo = new AppInfo;
        if(context.get_hostFullVersion())
            appInfo.hostVersion = context.get_hostFullVersion();
        else
            appInfo.hostVersion = context.get_appVersion();
        appInfo.appId = context.get_id();
        appInfo.host = context.get_appName();
        appInfo.browser = window.navigator.userAgent;
        appInfo.correlationId = trimStringToLowerCase(context.get_correlationId());
        appInfo.clientId = (new AppStorage).getClientId();
        appInfo.appInstanceId = context.get_appInstanceId();
        if(appInfo.appInstanceId)
            appInfo.appInstanceId = appInfo.appInstanceId.replace(/[{}]/g,"").toLowerCase();
        appInfo.message = context.get_hostCustomMessage();
        appInfo.officeJSVersion = OSF.ConstantNames.FileVersion;
        appInfo.hostJSVersion = "16.0.11527.30000";
        if(context._wacHostEnvironment)
            appInfo.wacHostEnvironment = context._wacHostEnvironment;
        if(context._isFromWacAutomation !== undefined && context._isFromWacAutomation !== null)
            appInfo.isFromWacAutomation = context._isFromWacAutomation.toString().toLowerCase();
        var docUrl = context.get_docUrl();
        appInfo.docUrl = omexDomainRegex.test(docUrl) ? docUrl : "";
        var url = location.href;
        if(url)
            url = url.split("?")[0].split("#")[0];
        appInfo.appURL = UrlFilter.filter(url);
        (function getUserIdAndAssetIdFromToken(token, appInfo)
        {
            var xmlContent;
            var parser;
            var xmlDoc;
            appInfo.assetId = "";
            appInfo.userId = "";
            try
            {
                xmlContent = decodeURIComponent(token);
                parser = new DOMParser;
                xmlDoc = parser.parseFromString(xmlContent,"text/xml");
                var cidNode = xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("cid");
                var oidNode = xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("oid");
                if(cidNode && cidNode.nodeValue)
                    appInfo.userId = cidNode.nodeValue;
                else if(oidNode && oidNode.nodeValue)
                    appInfo.userId = oidNode.nodeValue;
                appInfo.assetId = xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("aid").nodeValue
            }
            catch(e){}
            finally
            {
                xmlContent = null;
                xmlDoc = null;
                parser = null
            }
        })(context.get_eToken(),appInfo);
        appInfo.sessionId = sessionId;
        appInfo.name = context.get_addinName();
        if(typeof OTel !== "undefined")
            OTel.OTelLogger.initialize(appInfo);
        (function handleLifecycle()
        {
            var startTime = new Date;
            var lastFocus = null;
            var focusTime = 0;
            var finished = false;
            var adjustFocusTime = function()
                {
                    if(document.hasFocus())
                    {
                        if(lastFocus == null)
                            lastFocus = new Date
                    }
                    else if(lastFocus)
                    {
                        focusTime += Math.abs((new Date).getTime() - lastFocus.getTime());
                        lastFocus = null
                    }
                };
            var eventList = [];
            eventList.push(new Event("focus",adjustFocusTime));
            eventList.push(new Event("blur",adjustFocusTime));
            eventList.push(new Event("focusout",adjustFocusTime));
            eventList.push(new Event("focusin",adjustFocusTime));
            var exitFunction = function()
                {
                    for(var i = 0; i < eventList.length; i++)
                        OSF.OUtil.removeEventListener(window,eventList[i].name,eventList[i].handler);
                    eventList.length = 0;
                    if(!finished)
                    {
                        if(document.hasFocus() && lastFocus)
                        {
                            focusTime += Math.abs((new Date).getTime() - lastFocus.getTime());
                            lastFocus = null
                        }
                        OSFAppTelemetry.onAppClosed(Math.abs((new Date).getTime() - startTime.getTime()),focusTime);
                        finished = true
                    }
                };
            eventList.push(new Event("beforeunload",exitFunction));
            eventList.push(new Event("unload",exitFunction));
            for(var i = 0; i < eventList.length; i++)
                OSF.OUtil.addEventListener(window,eventList[i].name,eventList[i].handler);
            adjustFocusTime()
        })();
        OSFAppTelemetry.onAppActivated()
    }
    OSFAppTelemetry.initialize = initialize;
    function onAppActivated()
    {
        if(!appInfo)
            return;
        (new AppStorage).enumerateLog(function(id, log)
        {
            return(new AppLogger).LogRawData(log)
        },true);
        var data = new OSFLog.AppActivatedUsageData;
        data.SessionId = sessionId;
        data.AppId = appInfo.appId;
        data.AssetId = appInfo.assetId;
        data.AppURL = appInfo.appURL;
        data.UserId = "";
        data.ClientId = appInfo.clientId;
        data.Browser = appInfo.browser;
        data.Host = appInfo.host;
        data.HostVersion = appInfo.hostVersion;
        data.CorrelationId = trimStringToLowerCase(appInfo.correlationId);
        data.AppSizeWidth = window.innerWidth;
        data.AppSizeHeight = window.innerHeight;
        data.AppInstanceId = appInfo.appInstanceId;
        data.Message = appInfo.message;
        data.DocUrl = appInfo.docUrl;
        data.OfficeJSVersion = appInfo.officeJSVersion;
        data.HostJSVersion = appInfo.hostJSVersion;
        if(appInfo.wacHostEnvironment)
            data.WacHostEnvironment = appInfo.wacHostEnvironment;
        if(appInfo.isFromWacAutomation !== undefined && appInfo.isFromWacAutomation !== null)
            data.IsFromWacAutomation = appInfo.isFromWacAutomation;
        (new AppLogger).LogData(data)
    }
    OSFAppTelemetry.onAppActivated = onAppActivated;
    function onScriptDone(scriptId, msStartTime, msResponseTime, appCorrelationId)
    {
        var data = new OSFLog.ScriptLoadUsageData;
        data.CorrelationId = trimStringToLowerCase(appCorrelationId);
        data.SessionId = sessionId;
        data.ScriptId = scriptId;
        data.StartTime = msStartTime;
        data.ResponseTime = msResponseTime;
        (new AppLogger).LogData(data)
    }
    OSFAppTelemetry.onScriptDone = onScriptDone;
    function onCallDone(apiType, id, parameters, msResponseTime, errorType)
    {
        if(!appInfo)
            return;
        var data = new OSFLog.APIUsageUsageData;
        data.CorrelationId = trimStringToLowerCase(osfControlAppCorrelationId);
        data.SessionId = sessionId;
        data.APIType = apiType;
        data.APIID = id;
        data.Parameters = parameters;
        data.ResponseTime = msResponseTime;
        data.ErrorType = errorType;
        (new AppLogger).LogData(data)
    }
    OSFAppTelemetry.onCallDone = onCallDone;
    function onMethodDone(id, args, msResponseTime, errorType)
    {
        var parameters = null;
        if(args)
            if(typeof args == "number")
                parameters = String(args);
            else if(typeof args === "object")
                for(var index in args)
                {
                    if(parameters !== null)
                        parameters += ",";
                    else
                        parameters = "";
                    if(typeof args[index] == "number")
                        parameters += String(args[index])
                }
            else
                parameters = "";
        OSF.AppTelemetry.onCallDone("method",id,parameters,msResponseTime,errorType)
    }
    OSFAppTelemetry.onMethodDone = onMethodDone;
    function onPropertyDone(propertyName, msResponseTime)
    {
        OSF.AppTelemetry.onCallDone("property",-1,propertyName,msResponseTime)
    }
    OSFAppTelemetry.onPropertyDone = onPropertyDone;
    function onEventDone(id, errorType)
    {
        OSF.AppTelemetry.onCallDone("event",id,null,0,errorType)
    }
    OSFAppTelemetry.onEventDone = onEventDone;
    function onRegisterDone(register, id, msResponseTime, errorType)
    {
        OSF.AppTelemetry.onCallDone(register ? "registerevent" : "unregisterevent",id,null,msResponseTime,errorType)
    }
    OSFAppTelemetry.onRegisterDone = onRegisterDone;
    function onAppClosed(openTime, focusTime)
    {
        if(!appInfo)
            return;
        var data = new OSFLog.AppClosedUsageData;
        data.CorrelationId = trimStringToLowerCase(osfControlAppCorrelationId);
        data.SessionId = sessionId;
        data.FocusTime = focusTime;
        data.OpenTime = openTime;
        data.AppSizeFinalWidth = window.innerWidth;
        data.AppSizeFinalHeight = window.innerHeight;
        (new AppStorage).saveLog(sessionId,data.SerializeRow())
    }
    OSFAppTelemetry.onAppClosed = onAppClosed;
    function setOsfControlAppCorrelationId(correlationId)
    {
        osfControlAppCorrelationId = trimStringToLowerCase(correlationId)
    }
    OSFAppTelemetry.setOsfControlAppCorrelationId = setOsfControlAppCorrelationId;
    function doAppInitializationLogging(isException, message)
    {
        var data = new OSFLog.AppInitializationUsageData;
        data.CorrelationId = trimStringToLowerCase(osfControlAppCorrelationId);
        data.SessionId = sessionId;
        data.SuccessCode = isException ? 1 : 0;
        data.Message = message;
        (new AppLogger).LogData(data)
    }
    OSFAppTelemetry.doAppInitializationLogging = doAppInitializationLogging;
    function logAppCommonMessage(message)
    {
        doAppInitializationLogging(false,message)
    }
    OSFAppTelemetry.logAppCommonMessage = logAppCommonMessage;
    function logAppException(errorMessage)
    {
        doAppInitializationLogging(true,errorMessage)
    }
    OSFAppTelemetry.logAppException = logAppException;
    OSF.AppTelemetry = OSFAppTelemetry
})(OSFAppTelemetry || (OSFAppTelemetry = {}));
Microsoft.Office.WebExtension.EventType = {};
OSF.EventDispatch = function OSF_EventDispatch(eventTypes)
{
    this._eventHandlers = {};
    this._objectEventHandlers = {};
    this._queuedEventsArgs = {};
    if(eventTypes != null)
        for(var i = 0; i < eventTypes.length; i++)
        {
            var eventType = eventTypes[i];
            var isObjectEvent = eventType == "objectDeleted" || eventType == "objectSelectionChanged" || eventType == "objectDataChanged" || eventType == "contentControlAdded";
            if(!isObjectEvent)
                this._eventHandlers[eventType] = [];
            else
                this._objectEventHandlers[eventType] = {};
            this._queuedEventsArgs[eventType] = []
        }
};
OSF.EventDispatch.prototype = {
    getSupportedEvents: function OSF_EventDispatch$getSupportedEvents()
    {
        var events = [];
        for(var eventName in this._eventHandlers)
            events.push(eventName);
        for(var eventName in this._objectEventHandlers)
            events.push(eventName);
        return events
    },
    supportsEvent: function OSF_EventDispatch$supportsEvent(event)
    {
        for(var eventName in this._eventHandlers)
            if(event == eventName)
                return true;
        for(var eventName in this._objectEventHandlers)
            if(event == eventName)
                return true;
        return false
    },
    hasEventHandler: function OSF_EventDispatch$hasEventHandler(eventType, handler)
    {
        var handlers = this._eventHandlers[eventType];
        if(handlers && handlers.length > 0)
            for(var i = 0; i < handlers.length; i++)
                if(handlers[i] === handler)
                    return true;
        return false
    },
    hasObjectEventHandler: function OSF_EventDispatch$hasObjectEventHandler(eventType, objectId, handler)
    {
        var handlers = this._objectEventHandlers[eventType];
        if(handlers != null)
        {
            var _handlers = handlers[objectId];
            for(var i = 0; _handlers != null && i < _handlers.length; i++)
                if(_handlers[i] === handler)
                    return true
        }
        return false
    },
    addEventHandler: function OSF_EventDispatch$addEventHandler(eventType, handler)
    {
        if(typeof handler != "function")
            return false;
        var handlers = this._eventHandlers[eventType];
        if(handlers && !this.hasEventHandler(eventType,handler))
        {
            handlers.push(handler);
            return true
        }
        else
            return false
    },
    addObjectEventHandler: function OSF_EventDispatch$addObjectEventHandler(eventType, objectId, handler)
    {
        if(typeof handler != "function")
            return false;
        var handlers = this._objectEventHandlers[eventType];
        if(handlers && !this.hasObjectEventHandler(eventType,objectId,handler))
        {
            if(handlers[objectId] == null)
                handlers[objectId] = [];
            handlers[objectId].push(handler);
            return true
        }
        return false
    },
    addEventHandlerAndFireQueuedEvent: function OSF_EventDispatch$addEventHandlerAndFireQueuedEvent(eventType, handler)
    {
        var handlers = this._eventHandlers[eventType];
        var isFirstHandler = handlers.length == 0;
        var succeed = this.addEventHandler(eventType,handler);
        if(isFirstHandler && succeed)
            this.fireQueuedEvent(eventType);
        return succeed
    },
    removeEventHandler: function OSF_EventDispatch$removeEventHandler(eventType, handler)
    {
        var handlers = this._eventHandlers[eventType];
        if(handlers && handlers.length > 0)
            for(var index = 0; index < handlers.length; index++)
                if(handlers[index] === handler)
                {
                    handlers.splice(index,1);
                    return true
                }
        return false
    },
    removeObjectEventHandler: function OSF_EventDispatch$removeObjectEventHandler(eventType, objectId, handler)
    {
        var handlers = this._objectEventHandlers[eventType];
        if(handlers != null)
        {
            var _handlers = handlers[objectId];
            for(var i = 0; _handlers != null && i < _handlers.length; i++)
                if(_handlers[i] === handler)
                {
                    _handlers.splice(i,1);
                    return true
                }
        }
        return false
    },
    clearEventHandlers: function OSF_EventDispatch$clearEventHandlers(eventType)
    {
        if(typeof this._eventHandlers[eventType] != "undefined" && this._eventHandlers[eventType].length > 0)
        {
            this._eventHandlers[eventType] = [];
            return true
        }
        return false
    },
    clearObjectEventHandlers: function OSF_EventDispatch$clearObjectEventHandlers(eventType, objectId)
    {
        if(this._objectEventHandlers[eventType] != null && this._objectEventHandlers[eventType][objectId] != null)
        {
            this._objectEventHandlers[eventType][objectId] = [];
            return true
        }
        return false
    },
    getEventHandlerCount: function OSF_EventDispatch$getEventHandlerCount(eventType)
    {
        return this._eventHandlers[eventType] != undefined ? this._eventHandlers[eventType].length : -1
    },
    getObjectEventHandlerCount: function OSF_EventDispatch$getObjectEventHandlerCount(eventType, objectId)
    {
        if(this._objectEventHandlers[eventType] == null || this._objectEventHandlers[eventType][objectId] == null)
            return 0;
        return this._objectEventHandlers[eventType][objectId].length
    },
    fireEvent: function OSF_EventDispatch$fireEvent(eventArgs)
    {
        if(eventArgs.type == undefined)
            return false;
        var eventType = eventArgs.type;
        if(eventType && this._eventHandlers[eventType])
        {
            var eventHandlers = this._eventHandlers[eventType];
            for(var i = 0; i < eventHandlers.length; i++)
                eventHandlers[i](eventArgs);
            return true
        }
        else
            return false
    },
    fireObjectEvent: function OSF_EventDispatch$fireObjectEvent(objectId, eventArgs)
    {
        if(eventArgs.type == undefined)
            return false;
        var eventType = eventArgs.type;
        if(eventType && this._objectEventHandlers[eventType])
        {
            var eventHandlers = this._objectEventHandlers[eventType];
            var _handlers = eventHandlers[objectId];
            if(_handlers != null)
            {
                for(var i = 0; i < _handlers.length; i++)
                    _handlers[i](eventArgs);
                return true
            }
        }
        return false
    },
    fireOrQueueEvent: function OSF_EventDispatch$fireOrQueueEvent(eventArgs)
    {
        var eventType = eventArgs.type;
        if(eventType && this._eventHandlers[eventType])
        {
            var eventHandlers = this._eventHandlers[eventType];
            var queuedEvents = this._queuedEventsArgs[eventType];
            if(eventHandlers.length == 0)
                queuedEvents.push(eventArgs);
            else
                this.fireEvent(eventArgs);
            return true
        }
        else
            return false
    },
    fireQueuedEvent: function OSF_EventDispatch$queueEvent(eventType)
    {
        if(eventType && this._eventHandlers[eventType])
        {
            var eventHandlers = this._eventHandlers[eventType];
            var queuedEvents = this._queuedEventsArgs[eventType];
            if(eventHandlers.length > 0)
            {
                var eventHandler = eventHandlers[0];
                while(queuedEvents.length > 0)
                {
                    var eventArgs = queuedEvents.shift();
                    eventHandler(eventArgs)
                }
                return true
            }
        }
        return false
    },
    clearQueuedEvent: function OSF_EventDispatch$clearQueuedEvent(eventType)
    {
        if(eventType && this._eventHandlers[eventType])
        {
            var queuedEvents = this._queuedEventsArgs[eventType];
            if(queuedEvents)
                this._queuedEventsArgs[eventType] = []
        }
    }
};
OSF.DDA.OMFactory = OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureEventArgs = function OSF_DDA_OMFactory$manufactureEventArgs(eventType, target, eventProperties)
{
    var args;
    switch(eventType)
    {
        case Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged:
            args = new OSF.DDA.DocumentSelectionChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.BindingSelectionChanged:
            args = new OSF.DDA.BindingSelectionChangedEventArgs(this.manufactureBinding(eventProperties,target.document),eventProperties[OSF.DDA.PropertyDescriptors.Subset]);
            break;
        case Microsoft.Office.WebExtension.EventType.BindingDataChanged:
            args = new OSF.DDA.BindingDataChangedEventArgs(this.manufactureBinding(eventProperties,target.document));
            break;
        case Microsoft.Office.WebExtension.EventType.SettingsChanged:
            args = new OSF.DDA.SettingsChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.ActiveViewChanged:
            args = new OSF.DDA.ActiveViewChangedEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.OfficeThemeChanged:
            args = new OSF.DDA.Theming.OfficeThemeChangedEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.DocumentThemeChanged:
            args = new OSF.DDA.Theming.DocumentThemeChangedEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.AppCommandInvoked:
            args = OSF.DDA.AppCommand.AppCommandInvokedEventArgs.create(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.ObjectDeleted:
        case Microsoft.Office.WebExtension.EventType.ObjectSelectionChanged:
        case Microsoft.Office.WebExtension.EventType.ObjectDataChanged:
        case Microsoft.Office.WebExtension.EventType.ContentControlAdded:
            args = new OSF.DDA.ObjectEventArgs(eventType,eventProperties[Microsoft.Office.WebExtension.Parameters.Id]);
            break;
        case Microsoft.Office.WebExtension.EventType.RichApiMessage:
            args = new OSF.DDA.RichApiMessageEventArgs(eventType,eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.DataNodeInserted:
            args = new OSF.DDA.NodeInsertedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NewNode]),eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
            break;
        case Microsoft.Office.WebExtension.EventType.DataNodeReplaced:
            args = new OSF.DDA.NodeReplacedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.OldNode]),this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NewNode]),eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
            break;
        case Microsoft.Office.WebExtension.EventType.DataNodeDeleted:
            args = new OSF.DDA.NodeDeletedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.OldNode]),this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NextSiblingNode]),eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
            break;
        case Microsoft.Office.WebExtension.EventType.TaskSelectionChanged:
            args = new OSF.DDA.TaskSelectionChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.ResourceSelectionChanged:
            args = new OSF.DDA.ResourceSelectionChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.ViewSelectionChanged:
            args = new OSF.DDA.ViewSelectionChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.DialogMessageReceived:
            args = new OSF.DDA.DialogEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived:
            args = new OSF.DDA.DialogParentEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.ItemChanged:
            if(OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook")
            {
                args = new OSF.DDA.OlkItemSelectedChangedEventArgs(eventProperties);
                target.initialize(args["initialData"]);
                if(OSF._OfficeAppFactory.getHostInfo()["hostPlatform"] == "win32" || OSF._OfficeAppFactory.getHostInfo()["hostPlatform"] == "mac")
                    target.setCurrentItemNumber(args["itemNumber"].itemNumber)
            }
            else
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType,OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType,eventType));
            break;
        case Microsoft.Office.WebExtension.EventType.RecipientsChanged:
            if(OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook")
                args = new OSF.DDA.OlkRecipientsChangedEventArgs(eventProperties);
            else
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType,OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType,eventType));
            break;
        case Microsoft.Office.WebExtension.EventType.AppointmentTimeChanged:
            if(OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook")
                args = new OSF.DDA.OlkAppointmentTimeChangedEventArgs(eventProperties);
            else
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType,OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType,eventType));
            break;
        case Microsoft.Office.WebExtension.EventType.RecurrenceChanged:
            if(OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook")
                args = new OSF.DDA.OlkRecurrenceChangedEventArgs(eventProperties);
            else
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType,OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType,eventType));
            break;
        case Microsoft.Office.WebExtension.EventType.AttachmentsChanged:
            if(OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook")
                args = new OSF.DDA.OlkAttachmentsChangedEventArgs(eventProperties);
            else
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType,OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType,eventType));
            break;
        case Microsoft.Office.WebExtension.EventType.EnhancedLocationsChanged:
            if(OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook")
                args = new OSF.DDA.OlkEnhancedLocationsChangedEventArgs(eventProperties);
            else
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType,OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType,eventType));
            break;
        case Microsoft.Office.WebExtension.EventType.InfobarClicked:
            if(OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook")
                args = new OSF.DDA.OlkInfobarClickedEventArgs(eventProperties);
            else
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType,OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType,eventType));
            break;
        default:
            throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType,OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType,eventType));
    }
    return args
};
OSF.DDA.AsyncMethodNames.addNames({
    AddHandlerAsync: "addHandlerAsync",
    RemoveHandlerAsync: "removeHandlerAsync"
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.AddHandlerAsync,
    requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.EventType,
            "enum": Microsoft.Office.WebExtension.EventType,
            verify: function(eventType, caller, eventDispatch)
            {
                return eventDispatch.supportsEvent(eventType)
            }
        },{
            name: Microsoft.Office.WebExtension.Parameters.Handler,
            types: ["function"]
        }],
    supportedOptions: [],
    privateStateCallbacks: []
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.RemoveHandlerAsync,
    requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.EventType,
            "enum": Microsoft.Office.WebExtension.EventType,
            verify: function(eventType, caller, eventDispatch)
            {
                return eventDispatch.supportsEvent(eventType)
            }
        }],
    supportedOptions: [{
            name: Microsoft.Office.WebExtension.Parameters.Handler,
            value: {
                types: ["function","object"],
                defaultValue: null
            }
        }],
    privateStateCallbacks: []
});
var OfficeExt;
(function(OfficeExt)
{
    var AppCommand;
    (function(AppCommand)
    {
        var AppCommandManager = function()
            {
                function AppCommandManager()
                {
                    var _this = this;
                    this._pseudoDocument = null;
                    this._eventDispatch = null;
                    this._processAppCommandInvocation = function(args)
                    {
                        var verifyResult = _this._verifyManifestCallback(args.callbackName);
                        if(verifyResult.errorCode != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
                        {
                            _this._invokeAppCommandCompletedMethod(args.appCommandId,verifyResult.errorCode,"");
                            return
                        }
                        var eventObj = _this._constructEventObjectForCallback(args);
                        if(eventObj)
                            window.setTimeout(function()
                            {
                                verifyResult.callback(eventObj)
                            },0);
                        else
                            _this._invokeAppCommandCompletedMethod(args.appCommandId,OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError,"")
                    }
                }
                AppCommandManager.initializeOsfDda = function()
                {
                    OSF.DDA.AsyncMethodNames.addNames({AppCommandInvocationCompletedAsync: "appCommandInvocationCompletedAsync"});
                    OSF.DDA.AsyncMethodCalls.define({
                        method: OSF.DDA.AsyncMethodNames.AppCommandInvocationCompletedAsync,
                        requiredArguments: [{
                                name: Microsoft.Office.WebExtension.Parameters.Id,
                                types: ["string"]
                            },{
                                name: Microsoft.Office.WebExtension.Parameters.Status,
                                types: ["number"]
                            },{
                                name: Microsoft.Office.WebExtension.Parameters.AppCommandInvocationCompletedData,
                                types: ["string"]
                            }]
                    });
                    OSF.OUtil.augmentList(OSF.DDA.EventDescriptors,{AppCommandInvokedEvent: "AppCommandInvokedEvent"});
                    OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{AppCommandInvoked: "appCommandInvoked"});
                    OSF.OUtil.setNamespace("AppCommand",OSF.DDA);
                    OSF.DDA.AppCommand.AppCommandInvokedEventArgs = OfficeExt.AppCommand.AppCommandInvokedEventArgs
                };
                AppCommandManager.prototype.initializeAndChangeOnce = function(callback)
                {
                    AppCommand.registerDdaFacade();
                    this._pseudoDocument = {};
                    OSF.DDA.DispIdHost.addAsyncMethods(this._pseudoDocument,[OSF.DDA.AsyncMethodNames.AppCommandInvocationCompletedAsync,]);
                    this._eventDispatch = new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.AppCommandInvoked,]);
                    var onRegisterCompleted = function(result)
                        {
                            if(callback)
                                if(result.status == "succeeded")
                                    callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
                                else
                                    callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError)
                        };
                    OSF.DDA.DispIdHost.addEventSupport(this._pseudoDocument,this._eventDispatch);
                    this._pseudoDocument.addHandlerAsync(Microsoft.Office.WebExtension.EventType.AppCommandInvoked,this._processAppCommandInvocation,onRegisterCompleted)
                };
                AppCommandManager.prototype._verifyManifestCallback = function(callbackName)
                {
                    var defaultResult = {
                            callback: null,
                            errorCode: OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCallback
                        };
                    callbackName = callbackName.trim();
                    try
                    {
                        var callList = callbackName.split(".");
                        var parentObject = window;
                        for(var i = 0; i < callList.length - 1; i++)
                            if(parentObject[callList[i]] && (typeof parentObject[callList[i]] == "object" || typeof parentObject[callList[i]] == "function"))
                                parentObject = parentObject[callList[i]];
                            else
                                return defaultResult;
                        var callbackFunc = parentObject[callList[callList.length - 1]];
                        if(typeof callbackFunc != "function")
                            return defaultResult
                    }
                    catch(e)
                    {
                        return defaultResult
                    }
                    return{
                            callback: callbackFunc,
                            errorCode: OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess
                        }
                };
                AppCommandManager.prototype._invokeAppCommandCompletedMethod = function(appCommandId, resultCode, data)
                {
                    this._pseudoDocument.appCommandInvocationCompletedAsync(appCommandId,resultCode,data)
                };
                AppCommandManager.prototype._constructEventObjectForCallback = function(args)
                {
                    var _this = this;
                    var eventObj = new AppCommandCallbackEventArgs;
                    try
                    {
                        var jsonData = JSON.parse(args.eventObjStr);
                        this._translateEventObjectInternal(jsonData,eventObj);
                        Object.defineProperty(eventObj,"completed",{
                            value: function(completedContext)
                            {
                                eventObj.completedContext = completedContext;
                                var jsonString = JSON.stringify(eventObj);
                                _this._invokeAppCommandCompletedMethod(args.appCommandId,OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess,jsonString)
                            },
                            enumerable: true
                        })
                    }
                    catch(e)
                    {
                        eventObj = null
                    }
                    return eventObj
                };
                AppCommandManager.prototype._translateEventObjectInternal = function(input, output)
                {
                    for(var key in input)
                    {
                        if(!input.hasOwnProperty(key))
                            continue;
                        var inputChild = input[key];
                        if(typeof inputChild == "object" && inputChild != null)
                        {
                            OSF.OUtil.defineEnumerableProperty(output,key,{value: {}});
                            this._translateEventObjectInternal(inputChild,output[key])
                        }
                        else
                            Object.defineProperty(output,key,{
                                value: inputChild,
                                enumerable: true,
                                writable: true
                            })
                    }
                };
                AppCommandManager.prototype._constructObjectByTemplate = function(template, input)
                {
                    var output = {};
                    if(!template || !input)
                        return output;
                    for(var key in template)
                        if(template.hasOwnProperty(key))
                        {
                            output[key] = null;
                            if(input[key] != null)
                            {
                                var templateChild = template[key];
                                var inputChild = input[key];
                                var inputChildType = typeof inputChild;
                                if(typeof templateChild == "object" && templateChild != null)
                                    output[key] = this._constructObjectByTemplate(templateChild,inputChild);
                                else if(inputChildType == "number" || inputChildType == "string" || inputChildType == "boolean")
                                    output[key] = inputChild
                            }
                        }
                    return output
                };
                AppCommandManager.instance = function()
                {
                    if(AppCommandManager._instance == null)
                        AppCommandManager._instance = new AppCommandManager;
                    return AppCommandManager._instance
                };
                AppCommandManager._instance = null;
                return AppCommandManager
            }();
        AppCommand.AppCommandManager = AppCommandManager;
        var AppCommandInvokedEventArgs = function()
            {
                function AppCommandInvokedEventArgs(appCommandId, callbackName, eventObjStr)
                {
                    this.type = Microsoft.Office.WebExtension.EventType.AppCommandInvoked;
                    this.appCommandId = appCommandId;
                    this.callbackName = callbackName;
                    this.eventObjStr = eventObjStr
                }
                AppCommandInvokedEventArgs.create = function(eventProperties)
                {
                    return new AppCommandInvokedEventArgs(eventProperties[AppCommand.AppCommandInvokedEventEnums.AppCommandId],eventProperties[AppCommand.AppCommandInvokedEventEnums.CallbackName],eventProperties[AppCommand.AppCommandInvokedEventEnums.EventObjStr])
                };
                return AppCommandInvokedEventArgs
            }();
        AppCommand.AppCommandInvokedEventArgs = AppCommandInvokedEventArgs;
        var AppCommandCallbackEventArgs = function()
            {
                function AppCommandCallbackEventArgs(){}
                return AppCommandCallbackEventArgs
            }();
        AppCommand.AppCommandCallbackEventArgs = AppCommandCallbackEventArgs;
        AppCommand.AppCommandInvokedEventEnums = {
            AppCommandId: "appCommandId",
            CallbackName: "callbackName",
            EventObjStr: "eventObjStr"
        }
    })(AppCommand = OfficeExt.AppCommand || (OfficeExt.AppCommand = {}))
})(OfficeExt || (OfficeExt = {}));
OfficeExt.AppCommand.AppCommandManager.initializeOsfDda();
OSF.OUtil.setNamespace("Marshaling",OSF.DDA);
OSF.OUtil.setNamespace("AppCommand",OSF.DDA.Marshaling);
var OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys;
(function(OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys)
{
    OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys[OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys["AppCommandId"] = 0] = "AppCommandId";
    OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys[OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys["CallbackName"] = 1] = "CallbackName";
    OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys[OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys["EventObjStr"] = 2] = "EventObjStr"
})(OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys || (OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys = {}));
OSF.DDA.Marshaling.AppCommand.AppCommandInvokedEventKeys = OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys;
var OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys;
(function(OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys)
{
    OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys[OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys["Id"] = 0] = "Id";
    OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys[OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys["Status"] = 1] = "Status";
    OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys[OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys["Data"] = 2] = "Data"
})(OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys || (OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys = {}));
OSF.DDA.Marshaling.AppCommand.AppCommandCompletedMethodParameterKeys = OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys;
var OfficeExt;
(function(OfficeExt)
{
    var AppCommand;
    (function(AppCommand)
    {
        function registerDdaFacade()
        {
            if(OSF.DDA.WAC)
            {
                var parameterMap = OSF.DDA.WAC.Delegate.ParameterMap;
                parameterMap.define({
                    type: OSF.DDA.MethodDispId.dispidAppCommandInvocationCompletedMethod,
                    toHost: [{
                            name: Microsoft.Office.WebExtension.Parameters.Id,
                            value: OSF.DDA.Marshaling.AppCommand.AppCommandCompletedMethodParameterKeys.Id
                        },{
                            name: Microsoft.Office.WebExtension.Parameters.Status,
                            value: OSF.DDA.Marshaling.AppCommand.AppCommandCompletedMethodParameterKeys.Status
                        },{
                            name: Microsoft.Office.WebExtension.Parameters.AppCommandInvocationCompletedData,
                            value: OSF.DDA.Marshaling.AppCommand.AppCommandCompletedMethodParameterKeys.Data
                        }]
                });
                parameterMap.define({
                    type: OSF.DDA.EventDispId.dispidAppCommandInvokedEvent,
                    fromHost: [{
                            name: OSF.DDA.EventDescriptors.AppCommandInvokedEvent,
                            value: parameterMap.self
                        }]
                });
                parameterMap.addComplexType(OSF.DDA.EventDescriptors.AppCommandInvokedEvent);
                parameterMap.define({
                    type: OSF.DDA.EventDescriptors.AppCommandInvokedEvent,
                    fromHost: [{
                            name: OfficeExt.AppCommand.AppCommandInvokedEventEnums.AppCommandId,
                            value: OSF.DDA.Marshaling.AppCommand.AppCommandInvokedEventKeys.AppCommandId
                        },{
                            name: OfficeExt.AppCommand.AppCommandInvokedEventEnums.CallbackName,
                            value: OSF.DDA.Marshaling.AppCommand.AppCommandInvokedEventKeys.CallbackName
                        },{
                            name: OfficeExt.AppCommand.AppCommandInvokedEventEnums.EventObjStr,
                            value: OSF.DDA.Marshaling.AppCommand.AppCommandInvokedEventKeys.EventObjStr
                        },]
                })
            }
        }
        AppCommand.registerDdaFacade = registerDdaFacade
    })(AppCommand = OfficeExt.AppCommand || (OfficeExt.AppCommand = {}))
})(OfficeExt || (OfficeExt = {}));
OSF.DialogShownStatus = {
    hasDialogShown: false,
    isWindowDialog: false
};
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors,{DialogMessageReceivedEvent: "DialogMessageReceivedEvent"});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{
    DialogMessageReceived: "dialogMessageReceived",
    DialogEventReceived: "dialogEventReceived"
});
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors,{
    MessageType: "messageType",
    MessageContent: "messageContent"
});
OSF.DDA.DialogEventType = {};
OSF.OUtil.augmentList(OSF.DDA.DialogEventType,{
    DialogClosed: "dialogClosed",
    NavigationFailed: "naviationFailed"
});
OSF.DDA.AsyncMethodNames.addNames({
    DisplayDialogAsync: "displayDialogAsync",
    CloseAsync: "close"
});
OSF.DDA.SyncMethodNames.addNames({
    MessageParent: "messageParent",
    AddMessageHandler: "addEventHandler",
    SendMessage: "sendMessage"
});
OSF.DDA.UI.ParentUI = function OSF_DDA_ParentUI()
{
    var eventDispatch;
    if(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived != null)
        eventDispatch = new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DialogMessageReceived,Microsoft.Office.WebExtension.EventType.DialogEventReceived,Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived]);
    else
        eventDispatch = new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DialogMessageReceived,Microsoft.Office.WebExtension.EventType.DialogEventReceived]);
    var openDialogName = OSF.DDA.AsyncMethodNames.DisplayDialogAsync.displayName;
    var target = this;
    if(!target[openDialogName])
        OSF.OUtil.defineEnumerableProperty(target,openDialogName,{value: function()
            {
                var openDialog = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.OpenDialog];
                openDialog(arguments,eventDispatch,target)
            }});
    OSF.OUtil.finalizeProperties(this)
};
OSF.DDA.UI.ChildUI = function OSF_DDA_ChildUI(isPopupWindow)
{
    var messageParentName = OSF.DDA.SyncMethodNames.MessageParent.displayName;
    var target = this;
    if(!target[messageParentName])
        OSF.OUtil.defineEnumerableProperty(target,messageParentName,{value: function()
            {
                var messageParent = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.MessageParent];
                return messageParent(arguments,target)
            }});
    var addEventHandler = OSF.DDA.SyncMethodNames.AddMessageHandler.displayName;
    if(!target[addEventHandler] && typeof OSF.DialogParentMessageEventDispatch != "undefined")
        OSF.DDA.DispIdHost.addEventSupport(target,OSF.DialogParentMessageEventDispatch,isPopupWindow);
    OSF.OUtil.finalizeProperties(this)
};
OSF.DialogHandler = function OSF_DialogHandler(){};
OSF.DDA.DialogEventArgs = function OSF_DDA_DialogEventArgs(message)
{
    if(message[OSF.DDA.PropertyDescriptors.MessageType] == OSF.DialogMessageType.DialogMessageReceived)
        OSF.OUtil.defineEnumerableProperties(this,{
            type: {value: Microsoft.Office.WebExtension.EventType.DialogMessageReceived},
            message: {value: message[OSF.DDA.PropertyDescriptors.MessageContent]}
        });
    else
        OSF.OUtil.defineEnumerableProperties(this,{
            type: {value: Microsoft.Office.WebExtension.EventType.DialogEventReceived},
            error: {value: message[OSF.DDA.PropertyDescriptors.MessageType]}
        })
};
OSF.DDA.DialogParentEventArgs = function OSF_DDA_DialogParentEventArgs(message)
{
    OSF.OUtil.defineEnumerableProperties(this,{
        type: {value: Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived},
        message: {value: message[OSF.DDA.PropertyDescriptors.MessageContent]}
    })
};
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.DisplayDialogAsync,
    requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.Url,
            types: ["string"]
        }],
    supportedOptions: [{
            name: Microsoft.Office.WebExtension.Parameters.Width,
            value: {
                types: ["number"],
                defaultValue: 99
            }
        },{
            name: Microsoft.Office.WebExtension.Parameters.Height,
            value: {
                types: ["number"],
                defaultValue: 99
            }
        },{
            name: Microsoft.Office.WebExtension.Parameters.RequireHTTPs,
            value: {
                types: ["boolean"],
                defaultValue: true
            }
        },{
            name: Microsoft.Office.WebExtension.Parameters.DisplayInIframe,
            value: {
                types: ["boolean"],
                defaultValue: false
            }
        },{
            name: Microsoft.Office.WebExtension.Parameters.HideTitle,
            value: {
                types: ["boolean"],
                defaultValue: false
            }
        },{
            name: Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels,
            value: {
                types: ["boolean"],
                defaultValue: false
            }
        },{
            name: Microsoft.Office.WebExtension.Parameters.PromptBeforeOpen,
            value: {
                types: ["boolean"],
                defaultValue: true
            }
        },{
            name: Microsoft.Office.WebExtension.Parameters.EnforceAppDomain,
            value: {
                types: ["boolean"],
                defaultValue: false
            }
        }],
    privateStateCallbacks: [],
    onSucceeded: function(args, caller, callArgs)
    {
        var targetId = args[Microsoft.Office.WebExtension.Parameters.Id];
        var eventDispatch = args[Microsoft.Office.WebExtension.Parameters.Data];
        var dialog = new OSF.DialogHandler;
        var closeDialog = OSF.DDA.AsyncMethodNames.CloseAsync.displayName;
        OSF.OUtil.defineEnumerableProperty(dialog,closeDialog,{value: function()
            {
                var closeDialogfunction = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.CloseDialog];
                closeDialogfunction(arguments,targetId,eventDispatch,dialog)
            }});
        var addHandler = OSF.DDA.SyncMethodNames.AddMessageHandler.displayName;
        OSF.OUtil.defineEnumerableProperty(dialog,addHandler,{value: function()
            {
                var syncMethodCall = OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.AddMessageHandler.id];
                var callArgs = syncMethodCall.verifyAndExtractCall(arguments,dialog,eventDispatch);
                var eventType = callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
                var handler = callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
                return eventDispatch.addEventHandlerAndFireQueuedEvent(eventType,handler)
            }});
        var sendMessage = OSF.DDA.SyncMethodNames.SendMessage.displayName;
        OSF.OUtil.defineEnumerableProperty(dialog,sendMessage,{value: function()
            {
                var execute = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.SendMessage];
                return execute(arguments,eventDispatch,dialog)
            }});
        return dialog
    },
    checkCallArgs: function(callArgs, caller, stateInfo)
    {
        if(callArgs[Microsoft.Office.WebExtension.Parameters.Width] <= 0)
            callArgs[Microsoft.Office.WebExtension.Parameters.Width] = 1;
        if(!callArgs[Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels] && callArgs[Microsoft.Office.WebExtension.Parameters.Width] > 100)
            callArgs[Microsoft.Office.WebExtension.Parameters.Width] = 99;
        if(callArgs[Microsoft.Office.WebExtension.Parameters.Height] <= 0)
            callArgs[Microsoft.Office.WebExtension.Parameters.Height] = 1;
        if(!callArgs[Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels] && callArgs[Microsoft.Office.WebExtension.Parameters.Height] > 100)
            callArgs[Microsoft.Office.WebExtension.Parameters.Height] = 99;
        if(!callArgs[Microsoft.Office.WebExtension.Parameters.RequireHTTPs])
            callArgs[Microsoft.Office.WebExtension.Parameters.RequireHTTPs] = true;
        return callArgs
    }
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.CloseAsync,
    requiredArguments: [],
    supportedOptions: [],
    privateStateCallbacks: []
});
OSF.DDA.SyncMethodCalls.define({
    method: OSF.DDA.SyncMethodNames.MessageParent,
    requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.MessageToParent,
            types: ["string","number","boolean"]
        }],
    supportedOptions: []
});
OSF.DDA.SyncMethodCalls.define({
    method: OSF.DDA.SyncMethodNames.AddMessageHandler,
    requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.EventType,
            "enum": Microsoft.Office.WebExtension.EventType,
            verify: function(eventType, caller, eventDispatch)
            {
                return eventDispatch.supportsEvent(eventType)
            }
        },{
            name: Microsoft.Office.WebExtension.Parameters.Handler,
            types: ["function"]
        }],
    supportedOptions: []
});
OSF.DDA.SyncMethodCalls.define({
    method: OSF.DDA.SyncMethodNames.SendMessage,
    requiredArguments: [{
            name: Microsoft.Office.WebExtension.Parameters.MessageContent,
            types: ["string"]
        }],
    supportedOptions: [],
    privateStateCallbacks: []
});
OSF.OUtil.setNamespace("Marshaling",OSF.DDA);
OSF.OUtil.setNamespace("Dialog",OSF.DDA.Marshaling);
OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys = {
    MessageType: "messageType",
    MessageContent: "messageContent"
};
OSF.DDA.Marshaling.Dialog.DialogParentMessageReceivedEventKeys = {
    MessageType: "messageType",
    MessageContent: "messageContent"
};
OSF.DDA.Marshaling.MessageParentKeys = {MessageToParent: "messageToParent"};
OSF.DDA.Marshaling.DialogNotificationShownEventType = {DialogNotificationShown: "dialogNotificationShown"};
OSF.DDA.Marshaling.SendMessageKeys = {MessageContent: "messageContent"};
var OfficeExt;
(function(OfficeExt)
{
    var WacCommonUICssManager;
    (function(WacCommonUICssManager)
    {
        var hostType = {
                Excel: "excel",
                Word: "word",
                PowerPoint: "powerpoint",
                Outlook: "outlook",
                Visio: "visio"
            };
        function getDialogCssManager(applicationHostType)
        {
            switch(applicationHostType)
            {
                case hostType.Excel:
                case hostType.Word:
                case hostType.PowerPoint:
                case hostType.Outlook:
                case hostType.Visio:
                    return new DefaultDialogCSSManager;
                default:
                    return new DefaultDialogCSSManager
            }
            return null
        }
        WacCommonUICssManager.getDialogCssManager = getDialogCssManager;
        var DefaultDialogCSSManager = function()
            {
                function DefaultDialogCSSManager()
                {
                    this.overlayElementCSS = ["position: absolute","top: 0","left: 0","width: 100%","height: 100%","background-color: rgba(198, 198, 198, 0.5)","z-index: 99998"];
                    this.dialogNotificationPanelCSS = ["width: 100%","height: 190px","position: absolute","z-index: 99999","background-color: rgba(255, 255, 255, 1)","left: 0px","top: 50%","margin-top: -95px"];
                    this.newWindowNotificationTextPanelCSS = ["margin: 20px 14px","font-family: Segoe UI,Arial,Verdana,sans-serif","font-size: 14px","height: 100px","line-height: 100px"];
                    this.newWindowNotificationTextSpanCSS = ["display: inline-block","line-height: normal","vertical-align: middle"];
                    this.crossZoneNotificationTextPanelCSS = ["margin: 20px 14px","font-family: Segoe UI,Arial,Verdana,sans-serif","font-size: 14px","height: 100px",];
                    this.dialogNotificationButtonPanelCSS = "margin:0px 9px";
                    this.buttonStyleCSS = ["text-align: center","width: 70px","height: 25px","font-size: 14px","font-family: Segoe UI,Arial,Verdana,sans-serif","margin: 0px 5px","border-width: 1px","border-style: solid"]
                }
                DefaultDialogCSSManager.prototype.getOverlayElementCSS = function()
                {
                    return this.overlayElementCSS.join(";")
                };
                DefaultDialogCSSManager.prototype.getDialogNotificationPanelCSS = function()
                {
                    return this.dialogNotificationPanelCSS.join(";")
                };
                DefaultDialogCSSManager.prototype.getNewWindowNotificationTextPanelCSS = function()
                {
                    return this.newWindowNotificationTextPanelCSS.join(";")
                };
                DefaultDialogCSSManager.prototype.getNewWindowNotificationTextSpanCSS = function()
                {
                    return this.newWindowNotificationTextSpanCSS.join(";")
                };
                DefaultDialogCSSManager.prototype.getCrossZoneNotificationTextPanelCSS = function()
                {
                    return this.crossZoneNotificationTextPanelCSS.join(";")
                };
                DefaultDialogCSSManager.prototype.getDialogNotificationButtonPanelCSS = function()
                {
                    return this.dialogNotificationButtonPanelCSS
                };
                DefaultDialogCSSManager.prototype.getDialogButtonCSS = function()
                {
                    return this.buttonStyleCSS.join(";")
                };
                return DefaultDialogCSSManager
            }();
        WacCommonUICssManager.DefaultDialogCSSManager = DefaultDialogCSSManager
    })(WacCommonUICssManager = OfficeExt.WacCommonUICssManager || (OfficeExt.WacCommonUICssManager = {}))
})(OfficeExt || (OfficeExt = {}));
var OfficeExt;
(function(OfficeExt)
{
    var AddinNativeAction;
    (function(AddinNativeAction)
    {
        var Dialog;
        (function(Dialog)
        {
            var windowInstance = null;
            var handler = null;
            var overlayElement = null;
            var dialogNotificationPanel = null;
            var closeDialogKey = "osfDialogInternal:action=closeDialog";
            var showDialogCallback = null;
            var hasCrossZoneNotification = false;
            var checkWindowDialogCloseInterval = -1;
            var messageParentKey = "messageParentKey";
            var hostThemeButtonStyle = null;
            var commonButtonBorderColor = "#ababab";
            var commonButtonBackgroundColor = "#ffffff";
            var commonEventInButtonBackgroundColor = "#ccc";
            var newWindowNotificationId = "newWindowNotificaiton";
            var crossZoneNotificationId = "crossZoneNotification";
            var configureBrowserLinkId = "configureBrowserLink";
            var dialogNotificationTextPanelId = "dialogNotificationTextPanel";
            var shouldUseLocalStorageToPassMessage = OfficeExt.WACUtils.shouldUseLocalStorageToPassMessage();
            var registerDialogNotificationShownArgs = {
                    dispId: OSF.DDA.EventDispId.dispidDialogNotificationShownInAddinEvent,
                    eventType: OSF.DDA.Marshaling.DialogNotificationShownEventType.DialogNotificationShown,
                    onComplete: null
                };
            function setHostThemeButtonStyle(args)
            {
                var hostThemeButtonStyleArgs = args.input;
                if(hostThemeButtonStyleArgs != null)
                    hostThemeButtonStyle = {
                        HostButtonBorderColor: hostThemeButtonStyleArgs[OSF.HostThemeButtonStyleKeys.ButtonBorderColor],
                        HostButtonBackgroundColor: hostThemeButtonStyleArgs[OSF.HostThemeButtonStyleKeys.ButtonBackgroundColor]
                    };
                args.completed()
            }
            Dialog.setHostThemeButtonStyle = setHostThemeButtonStyle;
            function removeEventListenersForDialog(args)
            {
                OSF._OfficeAppFactory.getInitializationHelper().addOrRemoveEventListenersForWindow(false);
                args.completed()
            }
            Dialog.removeEventListenersForDialog = removeEventListenersForDialog;
            function handleNewWindowDialog(dialogInfo)
            {
                try
                {
                    if(dialogInfo[OSF.ShowWindowDialogParameterKeys.EnforceAppDomain] && !checkAppDomain(dialogInfo))
                    {
                        showDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeAppDomains);
                        return
                    }
                    if(!dialogInfo[OSF.ShowWindowDialogParameterKeys.PromptBeforeOpen])
                    {
                        showDialog(dialogInfo);
                        return
                    }
                    hasCrossZoneNotification = false;
                    var ignoreButtonKeyDownClick = false;
                    var hostInfoObj = OSF._OfficeAppFactory.getInitializationHelper()._hostInfo;
                    var dialogCssManager = OfficeExt.WacCommonUICssManager.getDialogCssManager(hostInfoObj.hostType);
                    var notificationText = OSF.OUtil.formatString(Strings.OfficeOM.L_ShowWindowDialogNotification,OSF._OfficeAppFactory.getInitializationHelper()._appContext._addinName);
                    overlayElement = createOverlayElement(dialogCssManager);
                    document.body.insertBefore(overlayElement,document.body.firstChild);
                    dialogNotificationPanel = createNotificationPanel(dialogCssManager,notificationText);
                    dialogNotificationPanel.id = newWindowNotificationId;
                    var dialogNotificationButtonPanel = createButtonPanel(dialogCssManager);
                    var allowButton = createButtonControl(dialogCssManager,Strings.OfficeOM.L_ShowWindowDialogNotificationAllow);
                    var ignoreButton = createButtonControl(dialogCssManager,Strings.OfficeOM.L_ShowWindowDialogNotificationIgnore);
                    dialogNotificationButtonPanel.appendChild(allowButton);
                    dialogNotificationButtonPanel.appendChild(ignoreButton);
                    dialogNotificationPanel.appendChild(dialogNotificationButtonPanel);
                    document.body.insertBefore(dialogNotificationPanel,document.body.firstChild);
                    allowButton.onclick = function(event)
                    {
                        showDialog(dialogInfo);
                        if(!hasCrossZoneNotification)
                            dismissDialogNotification();
                        event.preventDefault();
                        event.stopPropagation()
                    };
                    function ignoreButtonClickEventHandler(event)
                    {
                        function unregisterDialogNotificationShownEventCallback(status)
                        {
                            removeDialogNotificationElement();
                            setFocusOnFirstElement(status);
                            showDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeEndUserIgnore)
                        }
                        registerDialogNotificationShownArgs.onComplete = unregisterDialogNotificationShownEventCallback;
                        OSF.DDA.WAC.Delegate.unregisterEventAsync(registerDialogNotificationShownArgs);
                        event.preventDefault();
                        event.stopPropagation()
                    }
                    ignoreButton.onclick = ignoreButtonClickEventHandler;
                    allowButton.addEventListener("keydown",function(event)
                    {
                        if(event.shiftKey && event.keyCode == 9)
                        {
                            handleButtonControlEventOut(allowButton);
                            handleButtonControlEventIn(ignoreButton);
                            ignoreButton.focus();
                            event.preventDefault();
                            event.stopPropagation()
                        }
                    },false);
                    ignoreButton.addEventListener("keydown",function(event)
                    {
                        if(!event.shiftKey && event.keyCode == 9)
                        {
                            handleButtonControlEventOut(ignoreButton);
                            handleButtonControlEventIn(allowButton);
                            allowButton.focus();
                            event.preventDefault();
                            event.stopPropagation()
                        }
                        else if(event.keyCode == 13)
                        {
                            ignoreButtonKeyDownClick = true;
                            event.preventDefault();
                            event.stopPropagation()
                        }
                    },false);
                    ignoreButton.addEventListener("keyup",function(event)
                    {
                        if(event.keyCode == 13 && ignoreButtonKeyDownClick)
                        {
                            ignoreButtonKeyDownClick = false;
                            ignoreButtonClickEventHandler(event)
                        }
                    },false);
                    window.focus();
                    function registerDialogNotificationShownEventCallback(status)
                    {
                        allowButton.focus()
                    }
                    registerDialogNotificationShownArgs.onComplete = registerDialogNotificationShownEventCallback;
                    OSF.DDA.WAC.Delegate.registerEventAsync(registerDialogNotificationShownArgs)
                }
                catch(e)
                {
                    if(OSF.AppTelemetry)
                        OSF.AppTelemetry.logAppException("Exception happens at new window dialog." + e);
                    showDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError)
                }
            }
            Dialog.handleNewWindowDialog = handleNewWindowDialog;
            function closeDialog(callback)
            {
                try
                {
                    if(windowInstance != null)
                    {
                        var appDomains = OSF._OfficeAppFactory.getInitializationHelper()._appContext._appDomains;
                        if(appDomains)
                            for(var i = 0; i < appDomains.length && appDomains[i].indexOf("://") !== -1; i++)
                                windowInstance.postMessage(closeDialogKey,appDomains[i]);
                        if(windowInstance != null && !windowInstance.closed)
                            windowInstance.close();
                        if(shouldUseLocalStorageToPassMessage)
                            window.removeEventListener("storage",storageChangedHandler);
                        else
                            window.removeEventListener("message",receiveMessage);
                        window.clearInterval(checkWindowDialogCloseInterval);
                        windowInstance = null;
                        callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
                    }
                    else
                        callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError)
                }
                catch(e)
                {
                    if(OSF.AppTelemetry)
                        OSF.AppTelemetry.logAppException("Exception happens at close window dialog." + e);
                    callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError)
                }
            }
            Dialog.closeDialog = closeDialog;
            function messageParent(params)
            {
                var message = params.hostCallArgs[Microsoft.Office.WebExtension.Parameters.MessageToParent];
                if(shouldUseLocalStorageToPassMessage)
                    try
                    {
                        var messageKey = OSF._OfficeAppFactory.getInitializationHelper()._webAppState.id + messageParentKey;
                        window.localStorage.setItem(messageKey,message)
                    }
                    catch(e)
                    {
                        if(OSF.AppTelemetry)
                            OSF.AppTelemetry.logAppException("Error happened during messageParent method:" + e)
                    }
                else
                {
                    var appDomains = OSF._OfficeAppFactory.getInitializationHelper()._appContext._appDomains;
                    if(appDomains)
                        for(var i = 0; i < appDomains.length && appDomains[i].indexOf("://") !== -1; i++)
                            window.opener.postMessage(message,appDomains[i])
                }
            }
            Dialog.messageParent = messageParent;
            function sendMessage(params)
            {
                if(windowInstance != null)
                {
                    var message = params.hostCallArgs,
                        appDomains = OSF._OfficeAppFactory.getInitializationHelper()._appContext._appDomains;
                    if(appDomains)
                        for(var i = 0; i < appDomains.length && appDomains[i].indexOf("://") !== -1; i++)
                        {
                            if(typeof message != "string")
                                message = JSON.stringify(message);
                            windowInstance.postMessage(message,appDomains[i])
                        }
                }
            }
            Dialog.sendMessage = sendMessage;
            function registerMessageReceivedEvent()
            {
                function receiveCloseDialogMessage(event)
                {
                    if(event.source == window.opener)
                        if(typeof event.data === "string" && event.data.indexOf(closeDialogKey) > -1)
                            window.close();
                        else
                        {
                            var messageContent = event.data,
                                type = typeof messageContent;
                            if(messageContent && (type == "object" || type == "string"))
                            {
                                if(type == "string")
                                    messageContent = JSON.parse(messageContent);
                                var eventArgs = OSF.DDA.OMFactory.manufactureEventArgs(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived,null,messageContent);
                                OSF.DialogParentMessageEventDispatch.fireEvent(eventArgs)
                            }
                        }
                }
                window.addEventListener("message",receiveCloseDialogMessage)
            }
            Dialog.registerMessageReceivedEvent = registerMessageReceivedEvent;
            function setHandlerAndShowDialogCallback(onEventHandler, callback)
            {
                handler = onEventHandler;
                showDialogCallback = callback
            }
            Dialog.setHandlerAndShowDialogCallback = setHandlerAndShowDialogCallback;
            function escDismissDialogNotification()
            {
                try
                {
                    if(dialogNotificationPanel && dialogNotificationPanel.id == newWindowNotificationId && showDialogCallback)
                        showDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeEndUserIgnore)
                }
                catch(e)
                {
                    if(OSF.AppTelemetry)
                        OSF.AppTelemetry.logAppException("Error happened during executing displayDialogAsync callback." + e)
                }
                dismissDialogNotification()
            }
            Dialog.escDismissDialogNotification = escDismissDialogNotification;
            function showCrossZoneNotification(windowUrl, hostType)
            {
                var okButtonKeyDownClick = false;
                var dialogCssManager = OfficeExt.WacCommonUICssManager.getDialogCssManager(hostType);
                overlayElement = createOverlayElement(dialogCssManager);
                document.body.insertBefore(overlayElement,document.body.firstChild);
                dialogNotificationPanel = createNotificationPanelForCrossZoneIssue(dialogCssManager,windowUrl);
                dialogNotificationPanel.id = crossZoneNotificationId;
                var dialogNotificationButtonPanel = createButtonPanel(dialogCssManager);
                var okButton = createButtonControl(dialogCssManager,Strings.OfficeOM.L_DialogOK ? Strings.OfficeOM.L_DialogOK : "OK");
                dialogNotificationButtonPanel.appendChild(okButton);
                dialogNotificationPanel.appendChild(dialogNotificationButtonPanel);
                document.body.insertBefore(dialogNotificationPanel,document.body.firstChild);
                hasCrossZoneNotification = true;
                okButton.onclick = function()
                {
                    dismissDialogNotification()
                };
                okButton.addEventListener("keydown",function(event)
                {
                    if(event.keyCode == 9)
                    {
                        document.getElementById(configureBrowserLinkId).focus();
                        event.preventDefault();
                        event.stopPropagation()
                    }
                    else if(event.keyCode == 13)
                    {
                        okButtonKeyDownClick = true;
                        event.preventDefault();
                        event.stopPropagation()
                    }
                },false);
                okButton.addEventListener("keyup",function(event)
                {
                    if(event.keyCode == 13 && okButtonKeyDownClick)
                    {
                        okButtonKeyDownClick = false;
                        dismissDialogNotification();
                        event.preventDefault();
                        event.stopPropagation()
                    }
                },false);
                document.getElementById(configureBrowserLinkId).addEventListener("keydown",function(event)
                {
                    if(event.keyCode == 9)
                    {
                        okButton.focus();
                        event.preventDefault();
                        event.stopPropagation()
                    }
                },false);
                window.focus();
                okButton.focus()
            }
            Dialog.showCrossZoneNotification = showCrossZoneNotification;
            function receiveMessage(event)
            {
                if(event.source == windowInstance)
                    try
                    {
                        var dialogMessageReceivedArgs = {};
                        dialogMessageReceivedArgs[OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageType] = OSF.DialogMessageType.DialogMessageReceived;
                        dialogMessageReceivedArgs[OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageContent] = event.data;
                        handler(dialogMessageReceivedArgs)
                    }
                    catch(e)
                    {
                        if(OSF.AppTelemetry)
                            OSF.AppTelemetry.logAppException("Error happened during receive message handler." + e)
                    }
            }
            function storageChangedHandler(event)
            {
                var messageKey = OSF._OfficeAppFactory.getInitializationHelper()._webAppState.id + messageParentKey;
                if(event.key == messageKey)
                    try
                    {
                        var dialogMessageReceivedArgs = {};
                        dialogMessageReceivedArgs[OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageType] = OSF.DialogMessageType.DialogMessageReceived;
                        dialogMessageReceivedArgs[OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageContent] = event.newValue;
                        handler(dialogMessageReceivedArgs)
                    }
                    catch(e)
                    {
                        if(OSF.AppTelemetry)
                            OSF.AppTelemetry.logAppException("Error happened during storage changed handler." + e)
                    }
            }
            function checkAppDomain(dialogInfo)
            {
                var appDomains = OSF._OfficeAppFactory.getInitializationHelper()._appContext._appDomains;
                var url = dialogInfo[OSF.ShowWindowDialogParameterKeys.Url];
                return Microsoft.Office.Common.XdmCommunicationManager.checkUrlWithAppDomains(appDomains,url)
            }
            function showDialog(dialogInfo)
            {
                var hostInfoObj = OSF._OfficeAppFactory.getInitializationHelper()._hostInfo;
                var hostInfoVals = [hostInfoObj.hostType,hostInfoObj.hostPlatform,hostInfoObj.hostSpecificFileVersion,hostInfoObj.hostLocale,hostInfoObj.osfControlAppCorrelationId,"isDialog",hostInfoObj.disableLogging ? "disableLogging" : ""];
                var hostInfo = hostInfoVals.join("|");
                var appContext = OSF._OfficeAppFactory.getInitializationHelper()._appContext;
                var windowUrl = dialogInfo[OSF.ShowWindowDialogParameterKeys.Url];
                windowUrl = OfficeExt.WACUtils.addHostInfoAsQueryParam(windowUrl,hostInfo);
                var windowName = JSON.parse(window.name);
                windowName[OSF.WindowNameItemKeys.HostInfo] = hostInfo;
                windowName[OSF.WindowNameItemKeys.AppContext] = appContext;
                var width = dialogInfo[OSF.ShowWindowDialogParameterKeys.Width] * screen.width / 100;
                var height = dialogInfo[OSF.ShowWindowDialogParameterKeys.Height] * screen.height / 100;
                var left = appContext._clientWindowWidth / 2 - width / 2;
                var top = appContext._clientWindowHeight / 2 - height / 2;
                var windowSpecs = "width=" + width + ", height=" + height + ", left=" + left + ", top=" + top + ",channelmode=no,directories=no,fullscreen=no,location=no,menubar=no,resizable=yes,scrollbars=yes,status=no,titlebar=yes,toolbar=no";
                windowInstance = window.open(windowUrl,OfficeExt.WACUtils.serializeObjectToString(windowName),windowSpecs);
                if(windowInstance == null)
                {
                    OSF.AppTelemetry.logAppCommonMessage("Encountered cross zone issue in displayDialogAsync api.");
                    removeDialogNotificationElement();
                    showCrossZoneNotification(windowUrl,hostInfoObj.hostType);
                    showDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeCrossZone);
                    return
                }
                if(shouldUseLocalStorageToPassMessage)
                    window.addEventListener("storage",storageChangedHandler);
                else
                    window.addEventListener("message",receiveMessage);
                function checkWindowClose()
                {
                    try
                    {
                        if(windowInstance == null || windowInstance.closed)
                        {
                            window.clearInterval(checkWindowDialogCloseInterval);
                            if(shouldUseLocalStorageToPassMessage)
                                window.removeEventListener("storage",storageChangedHandler);
                            else
                                window.removeEventListener("message",receiveMessage);
                            var dialogClosedArgs = {};
                            dialogClosedArgs[OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageType] = OSF.DialogMessageType.DialogClosed;
                            handler(dialogClosedArgs)
                        }
                    }
                    catch(e)
                    {
                        if(OSF.AppTelemetry)
                            OSF.AppTelemetry.logAppException("Error happened during check or handle window close." + e)
                    }
                }
                checkWindowDialogCloseInterval = window.setInterval(checkWindowClose,1e3);
                if(showDialogCallback != null)
                    showDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
                else if(OSF.AppTelemetry)
                    OSF.AppTelemetry.logAppException("showDialogCallback can not be null.")
            }
            function createButtonControl(dialogCssManager, buttonValue)
            {
                var buttonControl = document.createElement("input");
                buttonControl.setAttribute("type","button");
                buttonControl.style.cssText = dialogCssManager.getDialogButtonCSS();
                buttonControl.style.borderColor = commonButtonBorderColor;
                buttonControl.style.backgroundColor = commonButtonBackgroundColor;
                buttonControl.setAttribute("value",buttonValue);
                var buttonControlEventInHandler = function()
                    {
                        handleButtonControlEventIn(buttonControl)
                    };
                var buttonControlEventOutHandler = function()
                    {
                        handleButtonControlEventOut(buttonControl)
                    };
                buttonControl.addEventListener("mouseover",buttonControlEventInHandler);
                buttonControl.addEventListener("focus",buttonControlEventInHandler);
                buttonControl.addEventListener("mouseout",buttonControlEventOutHandler);
                buttonControl.addEventListener("focusout",buttonControlEventOutHandler);
                return buttonControl
            }
            function handleButtonControlEventIn(buttonControl)
            {
                if(hostThemeButtonStyle != null)
                {
                    buttonControl.style.borderColor = hostThemeButtonStyle.HostButtonBorderColor;
                    buttonControl.style.backgroundColor = hostThemeButtonStyle.HostButtonBackgroundColor
                }
                else if(OSF.CommonUI && OSF.CommonUI.HostButtonBorderColor && OSF.CommonUI.HostButtonBackgroundColor)
                {
                    buttonControl.style.borderColor = OSF.CommonUI.HostButtonBorderColor;
                    buttonControl.style.backgroundColor = OSF.CommonUI.HostButtonBackgroundColor
                }
                else
                    buttonControl.style.backgroundColor = commonEventInButtonBackgroundColor
            }
            function handleButtonControlEventOut(buttonControl)
            {
                buttonControl.style.borderColor = commonButtonBorderColor;
                buttonControl.style.backgroundColor = commonButtonBackgroundColor
            }
            function dismissDialogNotification()
            {
                function unregisterDialogNotificationShownEventCallback(status)
                {
                    removeDialogNotificationElement();
                    setFocusOnFirstElement(status)
                }
                registerDialogNotificationShownArgs.onComplete = unregisterDialogNotificationShownEventCallback;
                OSF.DDA.WAC.Delegate.unregisterEventAsync(registerDialogNotificationShownArgs)
            }
            function removeDialogNotificationElement()
            {
                if(dialogNotificationPanel != null)
                {
                    document.body.removeChild(dialogNotificationPanel);
                    dialogNotificationPanel = null
                }
                if(overlayElement != null)
                {
                    document.body.removeChild(overlayElement);
                    overlayElement = null
                }
            }
            function createOverlayElement(dialogCssManager)
            {
                var overlayElement = document.createElement("div");
                overlayElement.style.cssText = dialogCssManager.getOverlayElementCSS();
                return overlayElement
            }
            function createNotificationPanel(dialogCssManager, notificationString)
            {
                var dialogNotificationPanel = document.createElement("div");
                dialogNotificationPanel.style.cssText = dialogCssManager.getDialogNotificationPanelCSS();
                setAttributeForDialogNotificationPanel(dialogNotificationPanel);
                var dialogNotificationTextPanel = document.createElement("div");
                dialogNotificationTextPanel.style.cssText = dialogCssManager.getNewWindowNotificationTextPanelCSS();
                dialogNotificationTextPanel.id = dialogNotificationTextPanelId;
                if(document.documentElement.getAttribute("dir") == "rtl")
                    dialogNotificationTextPanel.style.paddingRight = "30px";
                else
                    dialogNotificationTextPanel.style.paddingLeft = "30px";
                var dialogNotificationTextSpan = document.createElement("span");
                dialogNotificationTextSpan.style.cssText = dialogCssManager.getNewWindowNotificationTextSpanCSS();
                dialogNotificationTextSpan.innerText = notificationString;
                dialogNotificationTextPanel.appendChild(dialogNotificationTextSpan);
                dialogNotificationPanel.appendChild(dialogNotificationTextPanel);
                return dialogNotificationPanel
            }
            function createButtonPanel(dialogCssManager)
            {
                var dialogNotificationButtonPanel = document.createElement("div");
                dialogNotificationButtonPanel.style.cssText = dialogCssManager.getDialogNotificationButtonPanelCSS();
                if(document.documentElement.getAttribute("dir") == "rtl")
                    dialogNotificationButtonPanel.style.cssFloat = "left";
                else
                    dialogNotificationButtonPanel.style.cssFloat = "right";
                return dialogNotificationButtonPanel
            }
            function setFocusOnFirstElement(status)
            {
                if(status != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
                {
                    var list = document.querySelectorAll(OSF._OfficeAppFactory.getInitializationHelper()._tabbableElements);
                    OSF.OUtil.focusToFirstTabbable(list,false)
                }
            }
            function createNotificationPanelForCrossZoneIssue(dialogCssManager, windowUrl)
            {
                var dialogNotificationPanel = document.createElement("div");
                dialogNotificationPanel.style.cssText = dialogCssManager.getDialogNotificationPanelCSS();
                setAttributeForDialogNotificationPanel(dialogNotificationPanel);
                var dialogNotificationTextPanel = document.createElement("div");
                dialogNotificationTextPanel.style.cssText = dialogCssManager.getCrossZoneNotificationTextPanelCSS();
                dialogNotificationTextPanel.id = dialogNotificationTextPanelId;
                var configureBrowserLink = document.createElement("a");
                configureBrowserLink.id = configureBrowserLinkId;
                configureBrowserLink.href = "#";
                configureBrowserLink.innerText = Strings.OfficeOM.L_NewWindowCrossZoneConfigureBrowserLink;
                configureBrowserLink.setAttribute("onclick","window.open('https://support.microsoft.com/en-us/help/17479/windows-internet-explorer-11-change-security-privacy-settings', '_blank', 'fullscreen=1')");
                var dialogNotificationTextSpan = document.createElement("span");
                if(Strings.OfficeOM.L_NewWindowCrossZone)
                    dialogNotificationTextSpan.innerHTML = OSF.OUtil.formatString(Strings.OfficeOM.L_NewWindowCrossZone,configureBrowserLink.outerHTML,OfficeExt.WACUtils.getDomainForUrl(windowUrl));
                dialogNotificationTextPanel.appendChild(dialogNotificationTextSpan);
                dialogNotificationPanel.appendChild(dialogNotificationTextPanel);
                return dialogNotificationPanel
            }
            function setAttributeForDialogNotificationPanel(dialogNotificationDiv)
            {
                dialogNotificationDiv.setAttribute("role","dialog");
                dialogNotificationDiv.setAttribute("aria-describedby",dialogNotificationTextPanelId)
            }
        })(Dialog = AddinNativeAction.Dialog || (AddinNativeAction.Dialog = {}))
    })(AddinNativeAction = OfficeExt.AddinNativeAction || (OfficeExt.AddinNativeAction = {}))
})(OfficeExt || (OfficeExt = {}));
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidDialogMessageReceivedEvent,
    fromHost: [{
            name: OSF.DDA.EventDescriptors.DialogMessageReceivedEvent,
            value: OSF.DDA.WAC.Delegate.ParameterMap.self
        }]
});
OSF.DDA.WAC.Delegate.ParameterMap.addComplexType(OSF.DDA.EventDescriptors.DialogMessageReceivedEvent);
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDescriptors.DialogMessageReceivedEvent,
    fromHost: [{
            name: OSF.DDA.PropertyDescriptors.MessageType,
            value: OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageType
        },{
            name: OSF.DDA.PropertyDescriptors.MessageContent,
            value: OSF.DDA.Marshaling.Dialog.DialogMessageReceivedEventKeys.MessageContent
        }]
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidDialogParentMessageReceivedEvent,
    fromHost: [{
            name: OSF.DDA.EventDescriptors.DialogParentMessageReceivedEvent,
            value: OSF.DDA.WAC.Delegate.ParameterMap.self
        }]
});
OSF.DDA.WAC.Delegate.ParameterMap.addComplexType(OSF.DDA.EventDescriptors.DialogParentMessageReceivedEvent);
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDescriptors.DialogParentMessageReceivedEvent,
    fromHost: [{
            name: OSF.DDA.PropertyDescriptors.MessageType,
            value: OSF.DDA.Marshaling.Dialog.DialogParentMessageReceivedEventKeys.MessageType
        },{
            name: OSF.DDA.PropertyDescriptors.MessageContent,
            value: OSF.DDA.Marshaling.Dialog.DialogParentMessageReceivedEventKeys.MessageContent
        }]
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidMessageParentMethod,
    toHost: [{
            name: Microsoft.Office.WebExtension.Parameters.MessageToParent,
            value: OSF.DDA.Marshaling.MessageParentKeys.MessageToParent
        }]
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidSendMessageMethod,
    toHost: [{
            name: Microsoft.Office.WebExtension.Parameters.MessageContent,
            value: OSF.DDA.Marshaling.SendMessageKeys.MessageContent
        }]
});
OSF.DDA.WAC.Delegate.openDialog = function OSF_DDA_WAC_Delegate$OpenDialog(args)
{
    var httpsIdentifyString = "https://";
    var httpIdentifyString = "http://";
    var dialogInfo = JSON.parse(args.targetId);
    var callback = OSF.DDA.WAC.Delegate._getOnAfterRegisterEvent(true,args);
    function showDialogCallback(status)
    {
        var payload = {Error: status};
        try
        {
            callback(Microsoft.Office.Common.InvokeResultCode.noError,payload)
        }
        catch(e)
        {
            if(OSF.AppTelemetry)
                OSF.AppTelemetry.logAppException("Exception happens at showDialogCallback." + e)
        }
    }
    if(OSF.DialogShownStatus.hasDialogShown)
    {
        showDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened);
        return
    }
    var dialogUrl = dialogInfo[OSF.ShowWindowDialogParameterKeys.Url].toLowerCase();
    if(dialogUrl == null || !(dialogUrl.substr(0,httpsIdentifyString.length) === httpsIdentifyString))
    {
        if(dialogUrl.substr(0,httpIdentifyString.length) === httpIdentifyString)
            showDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeRequireHTTPS);
        else
            showDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidScheme);
        return
    }
    if(!dialogInfo[OSF.ShowWindowDialogParameterKeys.DisplayInIframe])
    {
        OSF.DialogShownStatus.isWindowDialog = true;
        OfficeExt.AddinNativeAction.Dialog.setHandlerAndShowDialogCallback(function OSF_DDA_WACDelegate$RegisterEventAsync_OnEvent(payload)
        {
            if(args.onEvent)
                args.onEvent(payload);
            if(OSF.AppTelemetry)
                OSF.AppTelemetry.onEventDone(args.dispId)
        },showDialogCallback);
        OfficeExt.AddinNativeAction.Dialog.handleNewWindowDialog(dialogInfo)
    }
    else
    {
        OSF.DialogShownStatus.isWindowDialog = false;
        OSF.DDA.WAC.Delegate.registerEventAsync(args)
    }
};
OSF.DDA.WAC.Delegate.messageParent = function OSF_DDA_WAC_Delegate$MessageParent(args)
{
    if(window.opener != null)
        OfficeExt.AddinNativeAction.Dialog.messageParent(args);
    else
        OSF.DDA.WAC.Delegate.executeAsync(args)
};
OSF.DDA.WAC.Delegate.sendMessage = function OSF_DDA_WAC_Delegate$SendMessage(args)
{
    if(OSF.DialogShownStatus.hasDialogShown)
        if(OSF.DialogShownStatus.isWindowDialog)
            OfficeExt.AddinNativeAction.Dialog.sendMessage(args);
        else
            OSF.DDA.WAC.Delegate.executeAsync(args)
};
OSF.DDA.WAC.Delegate.closeDialog = function OSF_DDA_WAC_Delegate$CloseDialog(args)
{
    var callback = OSF.DDA.WAC.Delegate._getOnAfterRegisterEvent(false,args);
    function closeDialogCallback(status)
    {
        var payload = {Error: status};
        try
        {
            callback(Microsoft.Office.Common.InvokeResultCode.noError,payload)
        }
        catch(e)
        {
            if(OSF.AppTelemetry)
                OSF.AppTelemetry.logAppException("Exception happens at closeDialogCallback." + e)
        }
    }
    if(!OSF.DialogShownStatus.hasDialogShown)
        closeDialogCallback(OSF.DDA.ErrorCodeManager.errorCodes.ooeWebDialogClosed);
    else if(OSF.DialogShownStatus.isWindowDialog)
    {
        if(args.onCalling)
            args.onCalling();
        OfficeExt.AddinNativeAction.Dialog.closeDialog(closeDialogCallback)
    }
    else
        OSF.DDA.WAC.Delegate.unregisterEventAsync(args)
};
OSF.InitializationHelper.prototype.dismissDialogNotification = function OSF_InitializationHelper$dismissDialogNotification()
{
    OfficeExt.AddinNativeAction.Dialog.escDismissDialogNotification()
};
OSF.InitializationHelper.prototype.registerMessageReceivedEventForWindowDialog = function OSF_InitializationHelper$registerMessageReceivedEventForWindowDialog()
{
    OfficeExt.AddinNativeAction.Dialog.registerMessageReceivedEvent()
};
OSF.DDA.AsyncMethodNames.addNames({CloseContainerAsync: "closeContainer"});
var OfficeExt;
(function(OfficeExt)
{
    var Container = function()
        {
            function Container(parameters){}
            return Container
        }();
    OfficeExt.Container = Container
})(OfficeExt || (OfficeExt = {}));
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.CloseContainerAsync,
    requiredArguments: [],
    supportedOptions: [],
    privateStateCallbacks: []
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidCloseContainerMethod,
    fromHost: [],
    toHost: []
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{ItemChanged: "olkItemSelectedChanged"});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors,{OlkItemSelectedData: "OlkItemSelectedData"});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{RecipientsChanged: "olkRecipientsChanged"});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors,{OlkRecipientsData: "OlkRecipientsData"});
OSF.DDA.OlkRecipientsChangedEventArgs = function OSF_DDA_OlkRecipientsChangedEventArgs(eventData)
{
    var changedRecipientFields = eventData[OSF.DDA.EventDescriptors.OlkRecipientsData][0];
    if(changedRecipientFields === "")
        changedRecipientFields = null;
    OSF.OUtil.defineEnumerableProperties(this,{
        type: {value: Microsoft.Office.WebExtension.EventType.RecipientsChanged},
        changedRecipientFields: {value: JSON.parse(changedRecipientFields)}
    })
};
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{AppointmentTimeChanged: "olkAppointmentTimeChanged"});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors,{OlkAppointmentTimeChangedData: "OlkAppointmentTimeChangedData"});
OSF.DDA.OlkAppointmentTimeChangedEventArgs = function OSF_DDA_OlkAppointmentTimeChangedEventArgs(eventData)
{
    var appointmentTimeString = eventData[OSF.DDA.EventDescriptors.OlkAppointmentTimeChangedData][0];
    var start;
    var end;
    try
    {
        var appointmentTime = JSON.parse(appointmentTimeString);
        start = new Date(appointmentTime.start).toISOString();
        end = new Date(appointmentTime.end).toISOString()
    }
    catch(e)
    {
        start = null;
        end = null
    }
    OSF.OUtil.defineEnumerableProperties(this,{
        type: {value: Microsoft.Office.WebExtension.EventType.AppointmentTimeChanged},
        start: {value: start},
        end: {value: end}
    })
};
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{RecurrenceChanged: "olkRecurrenceChanged"});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors,{OlkRecurrenceData: "OlkRecurrenceData"});
OSF.DDA.OlkRecurrenceChangedEventArgs = function OSF_DDA_OlkRecurrenceChangedEventArgs(eventData)
{
    var recurrenceObject = null;
    try
    {
        var dataObject = JSON.parse(eventData[OSF.DDA.EventDescriptors.OlkRecurrenceChangedData][0]);
        if(dataObject.recurrence != null)
        {
            recurrenceObject = JSON.parse(dataObject.recurrence);
            recurrenceObject = Microsoft.Office.WebExtension.OutlookBase.SeriesTimeJsonConverter(recurrenceObject)
        }
    }
    catch(e)
    {
        recurrenceObject = null
    }
    OSF.OUtil.defineEnumerableProperties(this,{
        type: {value: Microsoft.Office.WebExtension.EventType.RecurrenceChanged},
        recurrence: {value: recurrenceObject}
    })
};
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{OfficeThemeChanged: "officeThemeChanged"});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors,{OfficeThemeData: "OfficeThemeData"});
OSF.OUtil.setNamespace("Theming",OSF.DDA);
OSF.DDA.Theming.OfficeThemeChangedEventArgs = function OSF_DDA_Theming_OfficeThemeChangedEventArgs(officeTheme)
{
    var themeData = JSON.parse(officeTheme.OfficeThemeData[0]);
    var themeDataHex = {};
    for(var color in themeData)
        themeDataHex[color] = OSF.OUtil.convertIntToCssHexColor(themeData[color]);
    OSF.OUtil.defineEnumerableProperties(this,{
        type: {value: Microsoft.Office.WebExtension.EventType.OfficeThemeChanged},
        officeTheme: {value: themeDataHex}
    })
};
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{AttachmentsChanged: "olkAttachmentsChanged"});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors,{OlkAttachmentsChangedData: "OlkAttachmentsChangedData"});
OSF.DDA.OlkAttachmentsChangedEventArgs = function OSF_DDA_OlkAttachmentsChangedEventArgs(eventData)
{
    var attachmentStatus;
    var attachmentDetails;
    try
    {
        var attachmentChangedObject = JSON.parse(eventData[OSF.DDA.EventDescriptors.OlkAttachmentsChangedData][0]);
        attachmentStatus = attachmentChangedObject.attachmentStatus;
        attachmentDetails = Microsoft.Office.WebExtension.OutlookBase.CreateAttachmentDetails(attachmentChangedObject.attachmentDetails)
    }
    catch(e)
    {
        attachmentStatus = null;
        attachmentDetails = null
    }
    OSF.OUtil.defineEnumerableProperties(this,{
        type: {value: Microsoft.Office.WebExtension.EventType.AttachmentsChanged},
        attachmentStatus: {value: attachmentStatus},
        attachmentDetails: {value: attachmentDetails}
    })
};
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{EnhancedLocationsChanged: "olkEnhancedLocationsChanged"});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors,{OlkEnhancedLocationsChangedData: "OlkEnhancedLocationsChangedData"});
OSF.DDA.OlkEnhancedLocationsChangedEventArgs = function OSF_DDA_OlkEnhancedLocationsChangedEventArgs(eventData)
{
    var enhancedLocations;
    try
    {
        var enhancedLocationsChangedObject = JSON.parse(eventData[OSF.DDA.EventDescriptors.OlkEnhancedLocationsChangedData][0]);
        enhancedLocations = enhancedLocationsChangedObject.enhancedLocations
    }
    catch(e)
    {
        enhancedLocations = null
    }
    OSF.OUtil.defineEnumerableProperties(this,{
        type: {value: Microsoft.Office.WebExtension.EventType.EnhancedLocationsChanged},
        enhancedLocations: {value: enhancedLocations}
    })
};
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType,{InfobarClicked: "olkInfobarClicked"});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors,{OlkInfobarClickedData: "OlkInfobarClickedData"});
OSF.DDA.OlkInfobarClickedEventArgs = function OSF_DDA_OlkInfobarClickedEventArgs(eventData)
{
    var infobarDetails;
    try
    {
        infobarDetails = eventData[OSF.DDA.EventDescriptors.OlkInfobarClickedData][0]
    }
    catch(e)
    {
        infobarDetails = null
    }
    OSF.OUtil.defineEnumerableProperties(this,{
        type: {value: Microsoft.Office.WebExtension.EventType.InfobarClicked},
        infobarDetails: {value: infobarDetails}
    })
};
OSF.DDA.OlkItemSelectedChangedEventArgs = function OSF_DDA_OlkItemSelectedChangedEventArgs(eventData)
{
    var initialDataSource = eventData[OSF.DDA.EventDescriptors.OlkItemSelectedData];
    if(initialDataSource === "")
        initialDataSource = null;
    OSF.OUtil.defineEnumerableProperties(this,{
        type: {value: Microsoft.Office.WebExtension.EventType.ItemChanged},
        initialData: {value: initialDataSource}
    })
};
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkItemSelectedChangedEvent,
    fromHost: [{
            name: OSF.DDA.EventDescriptors.OlkItemSelectedData,
            value: OSF.DDA.WAC.Delegate.ParameterMap.self
        }],
    isComplexType: true
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkRecipientsChangedEvent,
    fromHost: [{
            name: OSF.DDA.EventDescriptors.OlkRecipientsData,
            value: OSF.DDA.WAC.Delegate.ParameterMap.self
        }],
    isComplexType: true
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkAppointmentTimeChangedEvent,
    fromHost: [{
            name: OSF.DDA.EventDescriptors.OlkAppointmentTimeChangedData,
            value: OSF.DDA.WAC.Delegate.ParameterMap.self
        }],
    isComplexType: true
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkRecurrenceChangedEvent,
    fromHost: [{
            name: OSF.DDA.EventDescriptors.OlkRecurrenceChangedData,
            value: OSF.DDA.WAC.Delegate.ParameterMap.self
        }],
    isComplexType: true
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOfficeThemeChangedEvent,
    fromHost: [{
            name: OSF.DDA.EventDescriptors.OfficeThemeData,
            value: OSF.DDA.WAC.Delegate.ParameterMap.self
        }],
    isComplexType: true
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkAttachmentsChangedEvent,
    fromHost: [{
            name: OSF.DDA.EventDescriptors.OlkAttachmentsChangedData,
            value: OSF.DDA.WAC.Delegate.ParameterMap.self
        }],
    isComplexType: true
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkEnhancedLocationsChangedEvent,
    fromHost: [{
            name: OSF.DDA.EventDescriptors.OlkEnhancedLocationsChangedData,
            value: OSF.DDA.WAC.Delegate.ParameterMap.self
        }],
    isComplexType: true
});
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkInfobarClickedEvent,
    fromHost: [{
            name: OSF.DDA.EventDescriptors.OlkInfobarClickedData,
            value: OSF.DDA.WAC.Delegate.ParameterMap.self
        }],
    isComplexType: true
});
OSF.DDA.AsyncMethodNames.addNames({GetAccessTokenAsync: "getAccessTokenAsync"});
OSF.DDA.Auth = function OSF_DDA_Auth(){};
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.GetAccessTokenAsync,
    requiredArguments: [],
    supportedOptions: [{
            name: Microsoft.Office.WebExtension.Parameters.ForceConsent,
            value: {
                types: ["boolean"],
                defaultValue: false
            }
        },{
            name: Microsoft.Office.WebExtension.Parameters.ForceAddAccount,
            value: {
                types: ["boolean"],
                defaultValue: false
            }
        },{
            name: Microsoft.Office.WebExtension.Parameters.AuthChallenge,
            value: {
                types: ["string"],
                defaultValue: ""
            }
        }],
    onSucceeded: function(dataDescriptor, caller, callArgs)
    {
        var data = dataDescriptor[Microsoft.Office.WebExtension.Parameters.Data];
        return data
    }
});
OSF.OUtil.setNamespace("Marshaling",OSF.DDA);
OSF.OUtil.setNamespace("SingleSignOn",OSF.DDA.Marshaling);
OSF.DDA.Marshaling.SingleSignOn.GetAccessTokenKeys = {
    ForceConsent: "forceConsent",
    ForceAddAccount: "forceAddAccount",
    AuthChallenge: "authChallenge"
};
OSF.DDA.Marshaling.SingleSignOn.AccessTokenResultKeys = {AccessToken: "accessToken"};
OSF.DDA.WAC.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidGetAccessTokenMethod,
    toHost: [{
            name: Microsoft.Office.WebExtension.Parameters.ForceConsent,
            value: OSF.DDA.Marshaling.SingleSignOn.GetAccessTokenKeys.ForceConsent
        },{
            name: Microsoft.Office.WebExtension.Parameters.ForceAddAccount,
            value: OSF.DDA.Marshaling.SingleSignOn.GetAccessTokenKeys.ForceAddAccount
        },{
            name: Microsoft.Office.WebExtension.Parameters.AuthChallenge,
            value: OSF.DDA.Marshaling.SingleSignOn.GetAccessTokenKeys.AuthChallenge
        }],
    fromHost: [{
            name: Microsoft.Office.WebExtension.Parameters.Data,
            value: OSF.DDA.Marshaling.SingleSignOn.AccessTokenResultKeys.AccessToken
        }]
});
OSF.InitializationHelper.prototype.prepareRightBeforeWebExtensionInitialize = function OSF_InitializationHelper$prepareRightBeforeWebExtensionInitialize(appContext)
{
    OSF.WebApp._UpdateLinksForHostAndXdmInfo()
};
OSF.InitializationHelper.prototype.prepareRightAfterWebExtensionInitialize = function()
{
    var appCommandHandler = OfficeExt.AppCommand.AppCommandManager.instance();
    appCommandHandler.initializeAndChangeOnce()
};
OSF.InitializationHelper.prototype.getInitializationReason = function OSF_InitializationHelper$getInitializationReason(appContext)
{
    return appContext.get_reason()
};
var executeAsyncBase = OSF.DDA.WAC.Delegate.executeAsync;
OSF.DDA.WAC.Delegate.executeAsync = function OSF_DDA_WAC_Delegate$executeAsyncOverride(args)
{
    var onCallingBase = args.onCalling;
    args.onCalling = function OSF_DDA_WAC_Delegate$executeAsync$onCalling()
    {
        args.hostCallArgs = OSF.DDA.OutlookAppOm.addAdditionalArgs(args.dispId,args.hostCallArgs);
        onCallingBase && onCallingBase()
    };
    executeAsyncBase(args)
};
OSF.InitializationHelper.prototype.prepareApiSurface = function OSF_InitializationHelper$prepareApiSurface(appContext)
{
    var license = new OSF.DDA.License(appContext.get_eToken());
    if(appContext.get_appName() == OSF.AppName.OutlookWebApp)
    {
        OSF.WebApp._UpdateLinksForHostAndXdmInfo();
        this.initWebDialog(appContext);
        this.initWebAuth(appContext);
        OSF._OfficeAppFactory.setContext(new OSF.DDA.OutlookContext(appContext,this._settings,license,appContext.appOM));
        OSF._OfficeAppFactory.setHostFacade(new OSF.DDA.DispIdHost.Facade(OSF.DDA.WAC.getDelegateMethods,OSF.DDA.WAC.Delegate.ParameterMap))
    }
    else
    {
        OfficeJsClient_OutlookWin32.prepareApiSurface(appContext);
        OSF._OfficeAppFactory.setContext(new OSF.DDA.OutlookContext(appContext,this._settings,license,appContext.appOM,OSF.DDA.OfficeTheme ? OSF.DDA.OfficeTheme.getOfficeTheme : null,appContext.ui));
        OSF._OfficeAppFactory.setHostFacade(new OSF.DDA.DispIdHost.Facade(OSF.DDA.DispIdHost.getClientDelegateMethods,OSF.DDA.SafeArray.Delegate.ParameterMap))
    }
};
OSF.DDA.SettingsManager = {
    SerializedSettings: "serializedSettings",
    DateJSONPrefix: "Date(",
    DataJSONSuffix: ")",
    serializeSettings: function OSF_DDA_SettingsManager$serializeSettings(settingsCollection)
    {
        var ret = {};
        for(var key in settingsCollection)
        {
            var value = settingsCollection[key];
            try
            {
                if(JSON)
                    value = JSON.stringify(value,function dateReplacer(k, v)
                    {
                        return OSF.OUtil.isDate(this[k]) ? OSF.DDA.SettingsManager.DateJSONPrefix + this[k].getTime() + OSF.DDA.SettingsManager.DataJSONSuffix : v
                    });
                else
                    value = Sys.Serialization.JavaScriptSerializer.serialize(value);
                ret[key] = value
            }
            catch(ex){}
        }
        return ret
    },
    deserializeSettings: function OSF_DDA_SettingsManager$deserializeSettings(serializedSettings)
    {
        var ret = {};
        serializedSettings = serializedSettings || {};
        for(var key in serializedSettings)
        {
            var value = serializedSettings[key];
            try
            {
                if(JSON)
                    value = JSON.parse(value,function dateReviver(k, v)
                    {
                        var d;
                        if(typeof v === "string" && v && v.length > 6 && v.slice(0,5) === OSF.DDA.SettingsManager.DateJSONPrefix && v.slice(-1) === OSF.DDA.SettingsManager.DataJSONSuffix)
                        {
                            d = new Date(parseInt(v.slice(5,-1)));
                            if(d)
                                return d
                        }
                        return v
                    });
                else
                    value = Sys.Serialization.JavaScriptSerializer.deserialize(value,true);
                ret[key] = value
            }
            catch(ex){}
        }
        return ret
    }
};
OSF.InitializationHelper.prototype.loadAppSpecificScriptAndCreateOM = function OSF_InitializationHelper$loadAppSpecificScriptAndCreateOM(appContext, appReady, basePath)
{
    Type.registerNamespace("Microsoft.Office.WebExtension.MailboxEnums");
    Microsoft.Office.WebExtension.MailboxEnums.EntityType = {
        MeetingSuggestion: "meetingSuggestion",
        TaskSuggestion: "taskSuggestion",
        Address: "address",
        EmailAddress: "emailAddress",
        Url: "url",
        PhoneNumber: "phoneNumber",
        Contact: "contact",
        FlightReservations: "flightReservations",
        ParcelDeliveries: "parcelDeliveries"
    };
    Microsoft.Office.WebExtension.MailboxEnums.ItemType = {
        Message: "message",
        Appointment: "appointment"
    };
    Microsoft.Office.WebExtension.MailboxEnums.ResponseType = {
        None: "none",
        Organizer: "organizer",
        Tentative: "tentative",
        Accepted: "accepted",
        Declined: "declined"
    };
    Microsoft.Office.WebExtension.MailboxEnums.RecipientType = {
        Other: "other",
        DistributionList: "distributionList",
        User: "user",
        ExternalUser: "externalUser"
    };
    Microsoft.Office.WebExtension.MailboxEnums.AttachmentType = {
        File: "file",
        Item: "item",
        Cloud: "cloud"
    };
    Microsoft.Office.WebExtension.MailboxEnums.AttachmentStatus = {
        Added: "added",
        Removed: "removed"
    };
    Microsoft.Office.WebExtension.MailboxEnums.AttachmentContentFormat = {
        Base64: "base64",
        Url: "url",
        Eml: "eml",
        ICalendar: "iCalendar"
    };
    Microsoft.Office.WebExtension.MailboxEnums.BodyType = {
        Text: "text",
        Html: "html"
    };
    Microsoft.Office.WebExtension.MailboxEnums.ItemNotificationMessageType = {
        ProgressIndicator: "progressIndicator",
        InformationalMessage: "informationalMessage",
        ErrorMessage: "errorMessage",
        InsightMessage: "insightMessage"
    };
    Microsoft.Office.WebExtension.MailboxEnums.Folder = {
        Inbox: "inbox",
        Junk: "junk",
        DeletedItems: "deletedItems"
    };
    Microsoft.Office.WebExtension.CoercionType = {
        Text: "text",
        Html: "html"
    };
    Microsoft.Office.WebExtension.MailboxEnums.UserProfileType = {
        Office365: "office365",
        OutlookCom: "outlookCom",
        Enterprise: "enterprise"
    };
    Microsoft.Office.WebExtension.MailboxEnums.RestVersion = {
        v1_0: "v1.0",
        v2_0: "v2.0",
        Beta: "beta"
    };
    Microsoft.Office.WebExtension.MailboxEnums.ModuleType = {Addins: "addins"};
    Microsoft.Office.WebExtension.MailboxEnums.ActionType = {ShowTaskPane: "showTaskPane"};
    Microsoft.Office.WebExtension.MailboxEnums.Days = {
        Mon: "mon",
        Tue: "tue",
        Wed: "wed",
        Thu: "thu",
        Fri: "fri",
        Sat: "sat",
        Sun: "sun",
        Weekday: "weekday",
        WeekendDay: "weekendDay",
        Day: "day"
    };
    Microsoft.Office.WebExtension.MailboxEnums.WeekNumber = {
        First: "first",
        Second: "second",
        Third: "third",
        Fourth: "fourth",
        Last: "last"
    };
    Microsoft.Office.WebExtension.MailboxEnums.RecurrenceType = {
        Daily: "daily",
        Weekday: "weekday",
        Weekly: "weekly",
        Monthly: "monthly",
        Yearly: "yearly"
    };
    Microsoft.Office.WebExtension.MailboxEnums.Month = {
        Jan: "jan",
        Feb: "feb",
        Mar: "mar",
        Apr: "apr",
        May: "may",
        Jun: "jun",
        Jul: "jul",
        Aug: "aug",
        Sep: "sep",
        Oct: "oct",
        Nov: "nov",
        Dec: "dec"
    };
    Microsoft.Office.WebExtension.MailboxEnums.DelegatePermissions = {
        Read: 1,
        Write: 2,
        DeleteOwn: 4,
        DeleteAll: 8,
        EditOwn: 16,
        EditAll: 32
    };
    Microsoft.Office.WebExtension.MailboxEnums.TimeZone = {
        AfghanistanStandardTime: "Afghanistan Standard Time",
        AlaskanStandardTime: "Alaskan Standard Time",
        AleutianStandardTime: "Aleutian Standard Time",
        AltaiStandardTime: "Altai Standard Time",
        ArabStandardTime: "Arab Standard Time",
        ArabianStandardTime: "Arabian Standard Time",
        ArabicStandardTime: "Arabic Standard Time",
        ArgentinaStandardTime: "Argentina Standard Time",
        AstrakhanStandardTime: "Astrakhan Standard Time",
        AtlanticStandardTime: "Atlantic Standard Time",
        AUSCentralStandardTime: "AUS Central Standard Time",
        AusCentralWStandardTime: "Aus Central W. Standard Time",
        AUSEasternStandardTime: "AUS Eastern Standard Time",
        AzerbaijanStandardTime: "Azerbaijan Standard Time",
        AzoresStandardTime: "Azores Standard Time",
        BahiaStandardTime: "Bahia Standard Time",
        BangladeshStandardTime: "Bangladesh Standard Time",
        BelarusStandardTime: "Belarus Standard Time",
        BougainvilleStandardTime: "Bougainville Standard Time",
        CanadaCentralStandardTime: "Canada Central Standard Time",
        CapeVerdeStandardTime: "Cape Verde Standard Time",
        CaucasusStandardTime: "Caucasus Standard Time",
        CenAustraliaStandardTime: "Cen. Australia Standard Time",
        CentralAmericaStandardTime: "Central America Standard Time",
        CentralAsiaStandardTime: "Central Asia Standard Time",
        CentralBrazilianStandardTime: "Central Brazilian Standard Time",
        CentralEuropeStandardTime: "Central Europe Standard Time",
        CentralEuropeanStandardTime: "Central European Standard Time",
        CentralPacificStandardTime: "Central Pacific Standard Time",
        CentralStandardTime: "Central Standard Time",
        CentralStandardTime_Mexico: "Central Standard Time (Mexico)",
        ChathamIslandsStandardTime: "Chatham Islands Standard Time",
        ChinaStandardTime: "China Standard Time",
        CubaStandardTime: "Cuba Standard Time",
        DatelineStandardTime: "Dateline Standard Time",
        EAfricaStandardTime: "E. Africa Standard Time",
        EAustraliaStandardTime: "E. Australia Standard Time",
        EEuropeStandardTime: "E. Europe Standard Time",
        ESouthAmericaStandardTime: "E. South America Standard Time",
        EasterIslandStandardTime: "Easter Island Standard Time",
        EasternStandardTime: "Eastern Standard Time",
        EasternStandardTime_Mexico: "Eastern Standard Time (Mexico)",
        EgyptStandardTime: "Egypt Standard Time",
        EkaterinburgStandardTime: "Ekaterinburg Standard Time",
        FijiStandardTime: "Fiji Standard Time",
        FLEStandardTime: "FLE Standard Time",
        GeorgianStandardTime: "Georgian Standard Time",
        GMTStandardTime: "GMT Standard Time",
        GreenlandStandardTime: "Greenland Standard Time",
        GreenwichStandardTime: "Greenwich Standard Time",
        GTBStandardTime: "GTB Standard Time",
        HaitiStandardTime: "Haiti Standard Time",
        HawaiianStandardTime: "Hawaiian Standard Time",
        IndiaStandardTime: "India Standard Time",
        IranStandardTime: "Iran Standard Time",
        IsraelStandardTime: "Israel Standard Time",
        JordanStandardTime: "Jordan Standard Time",
        KaliningradStandardTime: "Kaliningrad Standard Time",
        KamchatkaStandardTime: "Kamchatka Standard Time",
        KoreaStandardTime: "Korea Standard Time",
        LibyaStandardTime: "Libya Standard Time",
        LineIslandsStandardTime: "Line Islands Standard Time",
        LordHoweStandardTime: "Lord Howe Standard Time",
        MagadanStandardTime: "Magadan Standard Time",
        MagallanesStandardTime: "Magallanes Standard Time",
        MarquesasStandardTime: "Marquesas Standard Time",
        MauritiusStandardTime: "Mauritius Standard Time",
        MidAtlanticStandardTime: "Mid-Atlantic Standard Time",
        MiddleEastStandardTime: "Middle East Standard Time",
        MontevideoStandardTime: "Montevideo Standard Time",
        MoroccoStandardTime: "Morocco Standard Time",
        MountainStandardTime: "Mountain Standard Time",
        MountainStandardTime_Mexico: "Mountain Standard Time (Mexico)",
        MyanmarStandardTime: "Myanmar Standard Time",
        NCentralAsiaStandardTime: "N. Central Asia Standard Time",
        NamibiaStandardTime: "Namibia Standard Time",
        NepalStandardTime: "Nepal Standard Time",
        NewZealandStandardTime: "New Zealand Standard Time",
        NewfoundlandStandardTime: "Newfoundland Standard Time",
        NorfolkStandardTime: "Norfolk Standard Time",
        NorthAsiaEastStandardTime: "North Asia East Standard Time",
        NorthAsiaStandardTime: "North Asia Standard Time",
        NorthKoreaStandardTime: "North Korea Standard Time",
        OmskStandardTime: "Omsk Standard Time",
        PacificSAStandardTime: "Pacific SA Standard Time",
        PacificStandardTime: "Pacific Standard Time",
        PacificStandardTime_Mexico: "Pacific Standard Time (Mexico)",
        PakistanStandardTime: "Pakistan Standard Time",
        ParaguayStandardTime: "Paraguay Standard Time",
        RomanceStandardTime: "Romance Standard Time",
        RussiaTimeZone10: "Russia Time Zone 10",
        RussiaTimeZone11: "Russia Time Zone 11",
        RussiaTimeZone3: "Russia Time Zone 3",
        RussianStandardTime: "Russian Standard Time",
        SAEasternStandardTime: "SA Eastern Standard Time",
        SAPacificStandardTime: "SA Pacific Standard Time",
        SAWesternStandardTime: "SA Western Standard Time",
        SaintPierreStandardTime: "Saint Pierre Standard Time",
        SakhalinStandardTime: "Sakhalin Standard Time",
        SamoaStandardTime: "Samoa Standard Time",
        SaratovStandardTime: "Saratov Standard Time",
        SEAsiaStandardTime: "SE Asia Standard Time",
        SingaporeStandardTime: "Singapore Standard Time",
        SouthAfricaStandardTime: "South Africa Standard Time",
        SriLankaStandardTime: "Sri Lanka Standard Time",
        SudanStandardTime: "Sudan Standard Time",
        SyriaStandardTime: "Syria Standard Time",
        TaipeiStandardTime: "Taipei Standard Time",
        TasmaniaStandardTime: "Tasmania Standard Time",
        TocantinsStandardTime: "Tocantins Standard Time",
        TokyoStandardTime: "Tokyo Standard Time",
        TomskStandardTime: "Tomsk Standard Time",
        TongaStandardTime: "Tonga Standard Time",
        TransbaikalStandardTime: "Transbaikal Standard Time",
        TurkeyStandardTime: "Turkey Standard Time",
        TurksAndCaicosStandardTime: "Turks And Caicos Standard Time",
        UlaanbaatarStandardTime: "Ulaanbaatar Standard Time",
        USEasternStandardTime: "US Eastern Standard Time",
        USMountainStandardTime: "US Mountain Standard Time",
        UTC: "UTC",
        UTCPLUS12: "UTC+12",
        UTCPLUS13: "UTC+13",
        UTCMINUS02: "UTC-02",
        UTCMINUS08: "UTC-08",
        UTCMINUS09: "UTC-09",
        UTCMINUS11: "UTC-11",
        VenezuelaStandardTime: "Venezuela Standard Time",
        VladivostokStandardTime: "Vladivostok Standard Time",
        WAustraliaStandardTime: "W. Australia Standard Time",
        WCentralAfricaStandardTime: "W. Central Africa Standard Time",
        WEuropeStandardTime: "W. Europe Standard Time",
        WMongoliaStandardTime: "W. Mongolia Standard Time",
        WestAsiaStandardTime: "West Asia Standard Time",
        WestBankStandardTime: "West Bank Standard Time",
        WestPacificStandardTime: "West Pacific Standard Time",
        YakutskStandardTime: "Yakutsk Standard Time"
    };
    Microsoft.Office.WebExtension.MailboxEnums.LocationType = {
        Custom: "custom",
        Room: "room"
    };
    Microsoft.Office.WebExtension.MailboxEnums.CategoryColor = {
        None: "None",
        Preset0: "Preset0",
        Preset1: "Preset1",
        Preset2: "Preset2",
        Preset3: "Preset3",
        Preset4: "Preset4",
        Preset5: "Preset5",
        Preset6: "Preset6",
        Preset7: "Preset7",
        Preset8: "Preset8",
        Preset9: "Preset9",
        Preset10: "Preset10",
        Preset11: "Preset11",
        Preset12: "Preset12",
        Preset13: "Preset13",
        Preset14: "Preset14",
        Preset15: "Preset15",
        Preset16: "Preset16",
        Preset17: "Preset17",
        Preset18: "Preset18",
        Preset19: "Preset19",
        Preset20: "Preset20",
        Preset21: "Preset21",
        Preset22: "Preset22",
        Preset23: "Preset23",
        Preset24: "Preset24"
    };
    Type.registerNamespace("OSF.DDA");
    var OSF = window["OSF"] || {};
    OSF.DDA = OSF.DDA || {};
    window["OSF"]["DDA"]["OutlookAppOm"] = OSF.DDA.OutlookAppOm = function(officeAppContext, targetWindow, appReadyCallback)
    {
        this.$$d__getSharedPropertiesAsyncApi$p$0 = Function.createDelegate(this,this._getSharedPropertiesAsyncApi$p$0);
        this.$$d_navigateToModuleAsync = Function.createDelegate(this,this.navigateToModuleAsync);
        this.$$d_displayPersonaCardAsync = Function.createDelegate(this,this.displayPersonaCardAsync);
        this.$$d_displayNewMessageFormApi = Function.createDelegate(this,this.displayNewMessageFormApi);
        this.$$d__displayNewAppointmentFormApi$p$0 = Function.createDelegate(this,this._displayNewAppointmentFormApi$p$0);
        this.$$d_windowOpenOverrideHandler = Function.createDelegate(this,this.windowOpenOverrideHandler);
        this.$$d__getMasterCategories$p$0 = Function.createDelegate(this,this._getMasterCategories$p$0);
        this.$$d__getRestUrl$p$0 = Function.createDelegate(this,this._getRestUrl$p$0);
        this.$$d__getEwsUrl$p$0 = Function.createDelegate(this,this._getEwsUrl$p$0);
        this.$$d__getDiagnostics$p$0 = Function.createDelegate(this,this._getDiagnostics$p$0);
        this.$$d__getUserProfile$p$0 = Function.createDelegate(this,this._getUserProfile$p$0);
        this.$$d_getItem = Function.createDelegate(this,this.getItem);
        this.$$d__callAppReadyCallback$p$0 = Function.createDelegate(this,this._callAppReadyCallback$p$0);
        this.$$d__getInitialDataResponseHandler$p$0 = Function.createDelegate(this,this._getInitialDataResponseHandler$p$0);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p = this;
        this._officeAppContext$p$0 = officeAppContext;
        this._appReadyCallback$p$0 = appReadyCallback;
        var $$t_4 = this;
        var stringLoadedCallback = function()
            {
                if(appReadyCallback)
                    if(!$$t_4._officeAppContext$p$0["get_isDialog"]())
                        $$t_4.invokeHostMethod(1,null,$$t_4.$$d__getInitialDataResponseHandler$p$0);
                    else
                        window.setTimeout($$t_4.$$d__callAppReadyCallback$p$0,0)
            };
        if(this._areStringsLoaded$p$0())
            stringLoadedCallback();
        else
            this._loadLocalizedScript$p$0(stringLoadedCallback)
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i = function(currentPermissionLevel)
    {
        if(!currentPermissionLevel)
            throw Error.create(window["_u"]["ExtensibilityStrings"]["l_ElevatedPermissionNeeded_Text"]);
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i = function(value, minValue, maxValue, argumentName)
    {
        if(value < minValue || value > maxValue)
            throw Error.argumentOutOfRange(argumentName);
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidModule$p = function(module)
    {
        if($h.ScriptHelpers.isNullOrUndefined(module))
            throw Error.argumentNull("module");
        else if(module === "")
            throw Error.argument("module","module cannot be empty.");
        if(module !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ModuleType"]["Addins"])
            throw Error.notImplemented(String.format("API not supported for module '{0}'",module));
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._getHtmlBody$p = function(data)
    {
        var htmlBody = "";
        if("htmlBody" in data)
        {
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidHtmlBody$p(data["htmlBody"]);
            htmlBody = data["htmlBody"]
        }
        return htmlBody
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._getAttachments$p = function(data)
    {
        var attachments = [];
        if("attachments" in data)
        {
            attachments = data["attachments"];
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentsArray$p(attachments)
        }
        return attachments
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._getOptionsAndCallback$p = function(data)
    {
        var args = [];
        if("options" in data)
            args[0] = data["options"];
        if("callback" in data)
            args[args["length"]] = data["callback"];
        return args
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._createAttachmentsDataForHost$p = function(attachments)
    {
        var attachmentsData = new Array(0);
        if(Array["isInstanceOfType"](attachments))
            for(var i = 0; i < attachments["length"]; i++)
                if(Object["isInstanceOfType"](attachments[i]))
                {
                    var attachment = attachments[i];
                    window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachment$p(attachment);
                    attachmentsData[i] = window["OSF"]["DDA"]["OutlookAppOm"]._createAttachmentData$p(attachment)
                }
                else
                    throw Error.argument("attachments");
        return attachmentsData
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidHtmlBody$p = function(htmlBody)
    {
        if(!String["isInstanceOfType"](htmlBody))
            throw Error.argument("htmlBody");
        if($h.ScriptHelpers.isNullOrUndefined(htmlBody))
            throw Error.argument("htmlBody");
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(htmlBody.length,0,window["OSF"]["DDA"]["OutlookAppOm"].maxBodyLength,"htmlBody")
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentsArray$p = function(attachments)
    {
        if(!Array["isInstanceOfType"](attachments))
            throw Error.argument("attachments");
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachment$p = function(attachment)
    {
        if(!Object["isInstanceOfType"](attachment))
            throw Error.argument("attachments");
        if(!("type" in attachment) || !("name" in attachment))
            throw Error.argument("attachments");
        if(!("url" in attachment || "itemId" in attachment))
            throw Error.argument("attachments");
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._createAttachmentData$p = function(attachment)
    {
        var attachmentData = null;
        if(attachment["type"] === "file")
        {
            var url = attachment["url"];
            var name = attachment["name"];
            var isInline = $h.ScriptHelpers.isValueTrue(attachment["isInline"]);
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentUrlOrName$p(url,name);
            attachmentData = window["OSF"]["DDA"]["OutlookAppOm"]._createFileAttachmentData$p(url,name,isInline)
        }
        else if(attachment["type"] === "item")
        {
            var itemId = window["OSF"]["DDA"]["OutlookAppOm"].getItemIdBasedOnHost(attachment["itemId"]);
            var name = attachment["name"];
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentItemIdOrName$p(itemId,name);
            attachmentData = window["OSF"]["DDA"]["OutlookAppOm"]._createItemAttachmentData$p(itemId,name)
        }
        else
            throw Error.argument("attachments");
        return attachmentData
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._createFileAttachmentData$p = function(url, name, isInline)
    {
        return["file",name,url,isInline]
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._createItemAttachmentData$p = function(itemId, name)
    {
        return["item",name,itemId]
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentUrlOrName$p = function(url, name)
    {
        if(!String["isInstanceOfType"](url) || !String["isInstanceOfType"](name))
            throw Error.argument("attachments");
        if(url.length > 2048)
            throw Error.argumentOutOfRange("attachments",url.length,window["_u"]["ExtensibilityStrings"]["l_AttachmentUrlTooLong_Text"]);
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentName$p(name)
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentItemIdOrName$p = function(itemId, name)
    {
        if(!String["isInstanceOfType"](itemId) || !String["isInstanceOfType"](name))
            throw Error.argument("attachments");
        if(itemId.length > 200)
            throw Error.argumentOutOfRange("attachments",itemId.length,window["_u"]["ExtensibilityStrings"]["l_AttachmentItemIdTooLong_Text"]);
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentName$p(name)
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentName$p = function(name)
    {
        if(name.length > 255)
            throw Error.argumentOutOfRange("attachments",name.length,window["_u"]["ExtensibilityStrings"]["l_AttachmentNameTooLong_Text"]);
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidRestVersion$p = function(restVersion)
    {
        if(!restVersion)
            throw Error.argumentNull("restVersion");
        if(restVersion !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RestVersion"]["v1_0"] && restVersion !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RestVersion"]["v2_0"] && restVersion !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RestVersion"]["Beta"])
            throw Error.argument("restVersion");
    };
    window["OSF"]["DDA"]["OutlookAppOm"].getItemIdBasedOnHost = function(itemId)
    {
        if(window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._initialData$p$0 && window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._initialData$p$0.get__isRestIdSupported$i$0())
            return window["OSF"]["DDA"]["OutlookAppOm"]._instance$p["convertToRestId"](itemId,window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RestVersion"]["v1_0"]);
        return window["OSF"]["DDA"]["OutlookAppOm"]._instance$p["convertToEwsId"](itemId,window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RestVersion"]["v1_0"])
    };
    window["OSF"]["DDA"]["OutlookAppOm"].getErrorForTelemetry = function(resultCode, responseDictionary)
    {
        if(!$h.ScriptHelpers.isNullOrUndefined(resultCode) && resultCode)
            return resultCode;
        if(!responseDictionary)
            return-900;
        if("error" in responseDictionary)
        {
            if(!responseDictionary["error"])
                return 0;
            if("errorCode" in responseDictionary)
                return responseDictionary["errorCode"];
            else
                return-901
        }
        if("wasProxySuccessful" in responseDictionary)
            return responseDictionary["wasProxySuccessful"] ? 0 : -902;
        if("wasSuccessful" in responseDictionary)
            return responseDictionary["wasSuccessful"] ? 0 : -903;
        return-904
    };
    window["OSF"]["DDA"]["OutlookAppOm"]["addAdditionalArgs"] = function(dispid, data)
    {
        return data
    };
    window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType = function(value, expectedType, argumentName)
    {
        if(Object["getType"](value) !== expectedType)
            throw Error.argumentType(argumentName);
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._validateOptionalStringParameter$p = function(value, minLength, maxLength, name)
    {
        if($h.ScriptHelpers.isNullOrUndefined(value))
            return;
        window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(value,String,name);
        var stringValue = value;
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(stringValue.length,minLength,maxLength,name)
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p = function(array)
    {
        return array ? array["join"](";") : null
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._convertComposeEmailDictionaryParameterForSetApi$p = function(recipients)
    {
        if(!recipients)
            return null;
        var results = new Array(recipients["length"]);
        for(var i = 0; i < recipients["length"]; i++)
            results[i] = [recipients[i]["address"],recipients[i]["name"]];
        return results
    };
    window["OSF"]["DDA"]["OutlookAppOm"]._validateAndNormalizeRecipientEmails$p = function(emailset, name)
    {
        if($h.ScriptHelpers.isNullOrUndefined(emailset))
            return null;
        window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(emailset,Array,name);
        var originalAttendees = emailset;
        var updatedAttendees = null;
        var normalizationNeeded = false;
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(originalAttendees["length"],0,window["OSF"]["DDA"]["OutlookAppOm"]._maxRecipients$p,String.format("{0}.length",name));
        for(var i = 0; i < originalAttendees["length"]; i++)
            if($h.EmailAddressDetails["isInstanceOfType"](originalAttendees[i]))
            {
                normalizationNeeded = true;
                break
            }
        if(normalizationNeeded)
            updatedAttendees = [];
        for(var i = 0; i < originalAttendees["length"]; i++)
            if(normalizationNeeded)
            {
                updatedAttendees[i] = $h.EmailAddressDetails["isInstanceOfType"](originalAttendees[i]) ? originalAttendees[i]["emailAddress"] : originalAttendees[i];
                window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(updatedAttendees[i],String,String.format("{0}[{1}]",name,i))
            }
            else
                window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(originalAttendees[i],String,String.format("{0}[{1}]",name,i));
        return updatedAttendees
    };
    OSF.DDA.OutlookAppOm.prototype = {
        _initialData$p$0: null,
        _item$p$0: null,
        _userProfile$p$0: null,
        _diagnostics$p$0: null,
        _masterCategories$p$0: null,
        _officeAppContext$p$0: null,
        _appReadyCallback$p$0: null,
        _clientEndPoint$p$0: null,
        _hostItemType$p$0: 0,
        _additionalOutlookParams$p$0: null,
        get_clientEndPoint: function()
        {
            if(!this._clientEndPoint$p$0)
                this._clientEndPoint$p$0 = OSF._OfficeAppFactory["getClientEndPoint"]();
            return this._clientEndPoint$p$0
        },
        set_clientEndPoint: function(value)
        {
            this._clientEndPoint$p$0 = value;
            return value
        },
        get_initialData: function()
        {
            return this._initialData$p$0
        },
        get__appName$i$0: function()
        {
            return this._officeAppContext$p$0["get_appName"]()
        },
        get_additionalOutlookParams: function()
        {
            return this._additionalOutlookParams$p$0
        },
        addEventSupport: function()
        {
            if(this._item$p$0)
                OSF.DDA.DispIdHost["addEventSupport"](this._item$p$0,new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType["RecipientsChanged"],Microsoft.Office.WebExtension.EventType["AppointmentTimeChanged"],Microsoft.Office.WebExtension.EventType["RecurrenceChanged"],Microsoft.Office.WebExtension.EventType["AttachmentsChanged"],Microsoft.Office.WebExtension.EventType["EnhancedLocationsChanged"],Microsoft.Office.WebExtension.EventType["InfobarClicked"]]))
        },
        windowOpenOverrideHandler: function(url, targetName, features, replace)
        {
            this.invokeHostMethod(403,{launchUrl: url},null)
        },
        createAsyncResult: function(value, errorCode, detailedErrorCode, userContext, errorMessage)
        {
            var initArgs = {};
            var errorArgs = null;
            initArgs[OSF.DDA.AsyncResultEnum.Properties["Value"]] = value;
            initArgs[OSF.DDA.AsyncResultEnum.Properties["Context"]] = userContext;
            if(0 !== errorCode)
            {
                errorArgs = {};
                var errorProperties = $h.OutlookErrorManager.getErrorArgs(detailedErrorCode);
                errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties["Name"]] = errorProperties["name"];
                errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties["Message"]] = !errorMessage ? errorProperties["message"] : errorMessage;
                errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties["Code"]] = detailedErrorCode
            }
            return new OSF.DDA.AsyncResult(initArgs,errorArgs)
        },
        _throwOnMethodCallForInsufficientPermission$i$0: function(requiredPermissionLevel, methodName)
        {
            if(this._initialData$p$0._permissionLevel$p$0 < requiredPermissionLevel)
                throw Error.create(String.format(window["_u"]["ExtensibilityStrings"]["l_ElevatedPermissionNeededForMethod_Text"],methodName));
        },
        _displayReplyForm$i$0: function(obj)
        {
            this._displayReplyFormHelper$p$0(obj,false)
        },
        _displayReplyAllForm$i$0: function(obj)
        {
            this._displayReplyFormHelper$p$0(obj,true)
        },
        setActionsDefinition: function(actionsDefinition)
        {
            this._additionalOutlookParams$p$0.setActionsDefinition(actionsDefinition)
        },
        get_itemNumber: function()
        {
            return this._additionalOutlookParams$p$0._itemNumber$p$0
        },
        get_actionsDefinition: function()
        {
            return this._additionalOutlookParams$p$0._actionsDefinition$p$0
        },
        _displayReplyFormHelper$p$0: function(obj, isReplyAll)
        {
            if(String["isInstanceOfType"](obj))
                this._doDisplayReplyForm$p$0(obj,isReplyAll);
            else if(Object["isInstanceOfType"](obj) && Object.getTypeName(obj) === "Object")
                this._doDisplayReplyFormWithAttachments$p$0(obj,isReplyAll);
            else
                throw Error.argumentType();
        },
        _doDisplayReplyForm$p$0: function(htmlBody, isReplyAll)
        {
            if(!$h.ScriptHelpers.isNullOrUndefined(htmlBody))
                window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(htmlBody.length,0,window["OSF"]["DDA"]["OutlookAppOm"].maxBodyLength,"htmlBody");
            this.invokeHostMethod(isReplyAll ? 11 : 10,{htmlBody: htmlBody},null)
        },
        _doDisplayReplyFormWithAttachments$p$0: function(data, isReplyAll)
        {
            var htmlBody = window["OSF"]["DDA"]["OutlookAppOm"]._getHtmlBody$p(data);
            var attachments = window["OSF"]["DDA"]["OutlookAppOm"]._getAttachments$p(data);
            var parameters = $h.CommonParameters.parse(window["OSF"]["DDA"]["OutlookAppOm"]._getOptionsAndCallback$p(data),false);
            var $$t_6 = this;
            this._standardInvokeHostMethod$i$0(isReplyAll ? 31 : 30,{
                htmlBody: htmlBody,
                attachments: window["OSF"]["DDA"]["OutlookAppOm"]._createAttachmentsDataForHost$p(attachments)
            },function(rawInput)
            {
                return rawInput
            },parameters._asyncContext$p$0,parameters._callback$p$0)
        },
        _standardInvokeHostMethod$i$0: function(dispid, data, format, userContext, callback)
        {
            var $$t_C = this;
            this.invokeHostMethod(dispid,data,function(resultCode, response)
            {
                if(callback)
                {
                    var asyncResult = null;
                    var wasSuccessful = true;
                    if(Object["isInstanceOfType"](response))
                    {
                        var responseDictionary = response;
                        if("wasSuccessful" in responseDictionary)
                            wasSuccessful = responseDictionary["wasSuccessful"];
                        if("error" in responseDictionary || "data" in responseDictionary || "errorCode" in responseDictionary)
                            if(!responseDictionary["error"])
                            {
                                var formattedData = format ? format(responseDictionary["data"]) : responseDictionary["data"];
                                asyncResult = $$t_C.createAsyncResult(formattedData,0,0,userContext,null)
                            }
                            else
                            {
                                var errorCode = responseDictionary["errorCode"];
                                asyncResult = $$t_C.createAsyncResult(null,1,errorCode,userContext,null)
                            }
                    }
                    if(!asyncResult && resultCode)
                        asyncResult = $$t_C.createAsyncResult(null,1,9002,userContext,null);
                    if(!asyncResult && !resultCode && !wasSuccessful)
                        asyncResult = $$t_C.createAsyncResult(null,1,5e3,userContext,null);
                    callback(asyncResult)
                }
            })
        },
        getItemNumberFromOutlookResponse: function(responseData)
        {
            var itemNumber = 0;
            if(responseData["length"] > 2)
            {
                var extraParameters = window["JSON"]["parse"](responseData[2]);
                if(Object["isInstanceOfType"](extraParameters))
                {
                    var extraParametersDictionary = extraParameters;
                    itemNumber = extraParametersDictionary["itemNumber"]
                }
            }
            return itemNumber
        },
        createDeserializedData: function(responseData, itemChanged)
        {
            var deserializedData = null;
            var returnValues = window["JSON"]["parse"](responseData[0]);
            if(Object["isInstanceOfType"](returnValues))
                deserializedData = this._createDeserializedDataWithDictionary$p$0(responseData,itemChanged);
            else if(Number["isInstanceOfType"](returnValues))
                deserializedData = this._createDeserializedDataWithInt$p$0(responseData,itemChanged);
            else
                throw Error.notImplemented("Return data type from host must be Dictionary or int");
            return deserializedData
        },
        _createDeserializedDataWithDictionary$p$0: function(responseData, itemChanged)
        {
            var deserializedData = window["JSON"]["parse"](responseData[0]);
            if(itemChanged)
            {
                deserializedData["error"] = true;
                deserializedData["errorCode"] = 9030
            }
            else if(responseData["length"] > 1 && responseData[1])
            {
                deserializedData["error"] = true;
                deserializedData["errorCode"] = responseData[1];
                if(responseData["length"] > 2)
                {
                    var diagnosticsData = window["JSON"]["parse"](responseData[2]);
                    deserializedData["diagnostics"] = diagnosticsData["Diagnostics"]
                }
            }
            else
                deserializedData["error"] = false;
            return deserializedData
        },
        _createDeserializedDataWithInt$p$0: function(responseData, itemChanged)
        {
            var deserializedData = {};
            deserializedData["error"] = true;
            deserializedData["errorCode"] = responseData[0];
            return deserializedData
        },
        invokeHostMethod: function(dispid, data, responseCallback)
        {
            var startTime = (new Date)["getTime"]();
            var $$t_9 = this;
            var invokeResponseCallback = function(resultCode, resultData)
                {
                    if(window["OSF"]["AppTelemetry"])
                    {
                        var detailedErrorCode = window["OSF"]["DDA"]["OutlookAppOm"].getErrorForTelemetry(resultCode,resultData);
                        window["OSF"]["AppTelemetry"]["onMethodDone"](dispid,null,Math["abs"]((new Date)["getTime"]() - startTime),detailedErrorCode)
                    }
                    if(responseCallback)
                        responseCallback(resultCode,resultData)
                };
            if(64 === this._officeAppContext$p$0["get_appName"]())
            {
                var args = {ApiParams: data};
                args["MethodData"] = {
                    ControlId: OSF._OfficeAppFactory["getId"](),
                    DispatchId: dispid
                };
                args = window["OSF"]["DDA"]["OutlookAppOm"]["addAdditionalArgs"](dispid,args);
                if(dispid === 1)
                    this.get_clientEndPoint()["invoke"]("GetInitialData",invokeResponseCallback,args);
                else
                    this.get_clientEndPoint()["invoke"]("ExecuteMethod",invokeResponseCallback,args)
            }
            else if(!this._isOwaOnlyMethod$p$0(dispid))
                this.callOutlookDispatcher(dispid,data,responseCallback,startTime);
            else if(responseCallback)
                responseCallback(-2,null)
        },
        callOutlookDispatcher: function(dispid, data, responseCallback, startTime)
        {
            var executeParameters = this.convertToOutlookParameters(dispid,data);
            var $$t_D = this;
            OSF.ClientHostController["execute"](dispid,executeParameters,function(nativeData, resultCode)
            {
                var deserializedData = null;
                var responseData = nativeData.toArray();
                if(responseData["length"] > 0)
                {
                    var itemNumberFromOutlookResponse = $$t_D.getItemNumberFromOutlookResponse(responseData);
                    var isValidItemNumber = itemNumberFromOutlookResponse > 0;
                    var itemChanged = isValidItemNumber && itemNumberFromOutlookResponse > $$t_D._additionalOutlookParams$p$0._itemNumber$p$0;
                    deserializedData = $$t_D.createDeserializedData(responseData,itemChanged)
                }
                else if(responseCallback)
                    throw Error.argumentNull("responseData","Unexpected null/empty data from host.");
                if(window["OSF"]["AppTelemetry"])
                {
                    var detailedErrorCode = window["OSF"]["DDA"]["OutlookAppOm"].getErrorForTelemetry(resultCode,deserializedData);
                    window["OSF"]["AppTelemetry"]["onMethodDone"](dispid,null,Math["abs"]((new Date)["getTime"]() - startTime),detailedErrorCode)
                }
                if(responseCallback)
                    responseCallback(resultCode,deserializedData)
            })
        },
        _dictionaryToDate$i$0: function(input)
        {
            var retValue = new Date(input["year"],input["month"],input["date"],input["hours"],input["minutes"],input["seconds"],!input["milliseconds"] ? 0 : input["milliseconds"]);
            if(window["isNaN"](retValue["getTime"]()))
                throw Error.format(window["_u"]["ExtensibilityStrings"]["l_InvalidDate_Text"]);
            return retValue
        },
        _dateToDictionary$i$0: function(input)
        {
            var retValue = {};
            retValue["month"] = input["getMonth"]();
            retValue["date"] = input["getDate"]();
            retValue["year"] = input["getFullYear"]();
            retValue["hours"] = input["getHours"]();
            retValue["minutes"] = input["getMinutes"]();
            retValue["seconds"] = input["getSeconds"]();
            retValue["milliseconds"] = input["getMilliseconds"]();
            return retValue
        },
        _isOwaOnlyMethod$p$0: function(dispId)
        {
            switch(dispId)
            {
                case 402:
                case 401:
                case 400:
                case 403:
                    return true;
                default:
                    return false
            }
        },
        shouldRunNewCode: function(functionFlagToExecute)
        {
            return(this._initialData$p$0.get__shouldRunNewCodeForFlags$i$0() & functionFlagToExecute) === functionFlagToExecute
        },
        isOutlook16OrGreater: function()
        {
            var hostVersion = this._initialData$p$0.get__hostVersion$i$0();
            var endIndex = 0;
            var majorVersionNumber = 0;
            if(hostVersion)
            {
                endIndex = hostVersion.indexOf(".");
                majorVersionNumber = window["parseInt"](hostVersion.substring(0,endIndex))
            }
            return majorVersionNumber >= 16
        },
        isApiVersionSupported: function(requirementSet)
        {
            var apiSupported = false;
            try
            {
                var requirementDict = window["JSON"]["parse"](this._officeAppContext$p$0["get_requirementMatrix"]());
                var hostApiVersion = requirementDict["Mailbox"];
                var hostApiVersionParts = hostApiVersion.split(".");
                var requirementSetParts = requirementSet.split(".");
                if(window["parseInt"](hostApiVersionParts[0]) > window["parseInt"](requirementSetParts[0]) || window["parseInt"](hostApiVersionParts[0]) === window["parseInt"](requirementSetParts[0]) && window["parseInt"](hostApiVersionParts[1]) >= window["parseInt"](requirementSetParts[1]))
                    apiSupported = true
            }
            catch($$e_6){}
            return apiSupported
        },
        convertToOutlookParameters: function(dispid, data)
        {
            var executeParameters = null;
            var optionalParameters = {};
            switch(dispid)
            {
                case 1:
                case 2:
                case 3:
                case 14:
                case 18:
                case 26:
                case 32:
                case 41:
                case 34:
                case 99:
                case 103:
                case 107:
                case 108:
                case 149:
                case 154:
                case 157:
                case 160:
                case 164:
                    break;
                case 12:
                    optionalParameters["isRest"] = data["isRest"];
                    break;
                case 4:
                    var jsonProperty = window["JSON"]["stringify"](data["customProperties"]);
                    executeParameters = [jsonProperty];
                    break;
                case 5:
                    executeParameters = [data["body"]];
                    break;
                case 8:
                case 9:
                    executeParameters = [data["itemId"]];
                    break;
                case 7:
                    executeParameters = [window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["requiredAttendees"]),window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["optionalAttendees"]),data["start"],data["end"],data["location"],window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["resources"]),data["subject"],data["body"]];
                    break;
                case 44:
                    executeParameters = [window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["toRecipients"]),window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["ccRecipients"]),window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["bccRecipients"]),data["subject"],data["htmlBody"],data["attachments"]];
                    break;
                case 43:
                    executeParameters = [data["ewsIdOrEmail"]];
                    break;
                case 45:
                    executeParameters = [data["module"],data["queryString"]];
                    break;
                case 40:
                    executeParameters = [data["extensionId"],data["consentState"]];
                    break;
                case 11:
                case 10:
                    executeParameters = [data["htmlBody"]];
                    break;
                case 31:
                case 30:
                    executeParameters = [data["htmlBody"],data["attachments"]];
                    break;
                case 23:
                case 13:
                case 38:
                case 29:
                    executeParameters = [data["data"],data["coercionType"]];
                    break;
                case 37:
                case 28:
                    executeParameters = [data["coercionType"]];
                    break;
                case 17:
                    executeParameters = [data["subject"]];
                    break;
                case 15:
                    executeParameters = [data["recipientField"]];
                    break;
                case 22:
                case 21:
                    executeParameters = [data["recipientField"],window["OSF"]["DDA"]["OutlookAppOm"]._convertComposeEmailDictionaryParameterForSetApi$p(data["recipientArray"])];
                    break;
                case 19:
                    executeParameters = [data["itemId"],data["name"]];
                    break;
                case 16:
                    executeParameters = [data["uri"],data["name"],data["isInline"]];
                    break;
                case 148:
                    executeParameters = [data["base64String"],data["name"],data["isInline"]];
                    break;
                case 20:
                    executeParameters = [data["attachmentIndex"]];
                    break;
                case 25:
                    executeParameters = [data["TimeProperty"],data["time"]];
                    break;
                case 24:
                    executeParameters = [data["TimeProperty"]];
                    break;
                case 27:
                    executeParameters = [data["location"]];
                    break;
                case 33:
                case 35:
                    executeParameters = [data["key"],data["type"],data["persistent"],data["message"],data["icon"]];
                    this._additionalOutlookParams$p$0.setActionsDefinition(data["actions"]);
                    break;
                case 36:
                    executeParameters = [data["key"]];
                    break;
                case 100:
                case 150:
                case 101:
                case 104:
                case 151:
                case 152:
                case 153:
                case 155:
                case 156:
                case 158:
                case 159:
                case 161:
                case 162:
                case 163:
                    optionalParameters = data;
                    break;
                default:
                    Sys.Debug.fail("Unexpected method dispid");
                    break
            }
            if(dispid !== 1)
            {
                var $$t_5;
                this._additionalOutlookParams$p$0.updateOutlookExecuteParameters($$t_5 = {val: executeParameters},optionalParameters),executeParameters = $$t_5["val"]
            }
            return executeParameters
        },
        _displayNewAppointmentFormApi$p$0: function(parameters)
        {
            var normalizedRequiredAttendees = window["OSF"]["DDA"]["OutlookAppOm"]._validateAndNormalizeRecipientEmails$p(parameters["requiredAttendees"],"requiredAttendees");
            var normalizedOptionalAttendees = window["OSF"]["DDA"]["OutlookAppOm"]._validateAndNormalizeRecipientEmails$p(parameters["optionalAttendees"],"optionalAttendees");
            window["OSF"]["DDA"]["OutlookAppOm"]._validateOptionalStringParameter$p(parameters["location"],0,window["OSF"]["DDA"]["OutlookAppOm"]._maxLocationLength$p,"location");
            window["OSF"]["DDA"]["OutlookAppOm"]._validateOptionalStringParameter$p(parameters["body"],0,window["OSF"]["DDA"]["OutlookAppOm"].maxBodyLength,"body");
            window["OSF"]["DDA"]["OutlookAppOm"]._validateOptionalStringParameter$p(parameters["subject"],0,window["OSF"]["DDA"]["OutlookAppOm"]._maxSubjectLength$p,"subject");
            if(!$h.ScriptHelpers.isNullOrUndefined(parameters["start"]))
            {
                window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(parameters["start"],Date,"start");
                var startDateTime = parameters["start"];
                parameters["start"] = startDateTime["getTime"]();
                if(!$h.ScriptHelpers.isNullOrUndefined(parameters["end"]))
                {
                    window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(parameters["end"],Date,"end");
                    var endDateTime = parameters["end"];
                    if(endDateTime < startDateTime)
                        throw Error.argumentOutOfRange("end",endDateTime,window["_u"]["ExtensibilityStrings"]["l_InvalidEventDates_Text"]);
                    parameters["end"] = endDateTime["getTime"]()
                }
            }
            var updatedParameters = null;
            if(normalizedRequiredAttendees || normalizedOptionalAttendees)
            {
                updatedParameters = {};
                var $$dict_7 = parameters;
                for(var $$key_8 in $$dict_7)
                {
                    var entry = {
                            key: $$key_8,
                            value: $$dict_7[$$key_8]
                        };
                    updatedParameters[entry["key"]] = entry["value"]
                }
                if(normalizedRequiredAttendees)
                    updatedParameters["requiredAttendees"] = normalizedRequiredAttendees;
                if(normalizedOptionalAttendees)
                    updatedParameters["optionalAttendees"] = normalizedOptionalAttendees
            }
            this.invokeHostMethod(7,updatedParameters || parameters,null)
        },
        displayNewMessageFormApi: function(parameters)
        {
            var updatedParameters = {};
            if(parameters)
            {
                var normalizedToRecipients = window["OSF"]["DDA"]["OutlookAppOm"]._validateAndNormalizeRecipientEmails$p(parameters["toRecipients"],"toRecipients");
                var normalizedCcRecipients = window["OSF"]["DDA"]["OutlookAppOm"]._validateAndNormalizeRecipientEmails$p(parameters["ccRecipients"],"ccRecipients");
                var normalizedBccRecipients = window["OSF"]["DDA"]["OutlookAppOm"]._validateAndNormalizeRecipientEmails$p(parameters["bccRecipients"],"bccRecipients");
                window["OSF"]["DDA"]["OutlookAppOm"]._validateOptionalStringParameter$p(parameters["htmlBody"],0,window["OSF"]["DDA"]["OutlookAppOm"].maxBodyLength,"htmlBody");
                window["OSF"]["DDA"]["OutlookAppOm"]._validateOptionalStringParameter$p(parameters["subject"],0,window["OSF"]["DDA"]["OutlookAppOm"]._maxSubjectLength$p,"subject");
                var attachments = window["OSF"]["DDA"]["OutlookAppOm"]._getAttachments$p(parameters);
                var $$dict_7 = parameters;
                for(var $$key_8 in $$dict_7)
                {
                    var entry = {
                            key: $$key_8,
                            value: $$dict_7[$$key_8]
                        };
                    updatedParameters[entry["key"]] = entry["value"]
                }
                if(normalizedToRecipients)
                    updatedParameters["toRecipients"] = normalizedToRecipients;
                if(normalizedCcRecipients)
                    updatedParameters["ccRecipients"] = normalizedCcRecipients;
                if(normalizedBccRecipients)
                    updatedParameters["bccRecipients"] = normalizedBccRecipients;
                if(attachments)
                    updatedParameters["attachments"] = window["OSF"]["DDA"]["OutlookAppOm"]._createAttachmentsDataForHost$p(attachments)
            }
            this.invokeHostMethod(44,updatedParameters || parameters,null)
        },
        displayPersonaCardAsync: function(ewsIdOrEmail)
        {
            var args = [];
            for(var $$pai_3 = 1; $$pai_3 < arguments["length"]; ++$$pai_3)
                args[$$pai_3 - 1] = arguments[$$pai_3];
            if($h.ScriptHelpers.isNullOrUndefined(ewsIdOrEmail))
                throw Error.argumentNull("ewsIdOrEmail");
            else if(ewsIdOrEmail === "")
                throw Error.argument("ewsIdOrEmail","ewsIdOrEmail cannot be empty.");
            var parameters = $h.CommonParameters.parse(args,false);
            this._standardInvokeHostMethod$i$0(43,{ewsIdOrEmail: ewsIdOrEmail.trim()},null,parameters._asyncContext$p$0,parameters._callback$p$0)
        },
        navigateToModuleAsync: function(module)
        {
            var args = [];
            for(var $$pai_5 = 1; $$pai_5 < arguments["length"]; ++$$pai_5)
                args[$$pai_5 - 1] = arguments[$$pai_5];
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidModule$p(module);
            var parameters = $h.CommonParameters.parse(args,false);
            var updatedParameters = {};
            if(module === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ModuleType"]["Addins"])
            {
                var queryString = "";
                if(parameters._options$p$0 && parameters._options$p$0["queryString"])
                    queryString = parameters._options$p$0["queryString"];
                updatedParameters["queryString"] = queryString
            }
            updatedParameters["module"] = module;
            this._standardInvokeHostMethod$i$0(45,updatedParameters,null,parameters._asyncContext$p$0,parameters._callback$p$0)
        },
        _getSharedPropertiesAsyncApi$p$0: function()
        {
            var args = [];
            for(var $$pai_2 = 0; $$pai_2 < arguments["length"]; ++$$pai_2)
                args[$$pai_2] = arguments[$$pai_2];
            window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"getSharedPropertiesAsync");
            var parameters = $h.CommonParameters.parse(args,true);
            window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(108,null,null,parameters._asyncContext$p$0,parameters._callback$p$0)
        },
        _initializeMethods$p$0: function()
        {
            var currentInstance = this;
            if($h.Item["isInstanceOfType"](this._item$p$0) || this._hostItemType$p$0 === 6)
            {
                currentInstance["displayNewAppointmentForm"] = this.$$d__displayNewAppointmentFormApi$p$0;
                currentInstance["displayNewMessageForm"] = this.$$d_displayNewMessageFormApi;
                currentInstance["displayPersonaCardAsync"] = this.$$d_displayPersonaCardAsync;
                currentInstance["navigateToModuleAsync"] = this.$$d_navigateToModuleAsync
            }
            if(this._item$p$0 && this._item$p$0.get__isFromSharedFolder$i$0() && this._hostItemType$p$0 !== 6)
                this._item$p$0["getSharedPropertiesAsync"] = this.$$d__getSharedPropertiesAsyncApi$p$0
        },
        _getInitialDataResponseHandler$p$0: function(resultCode, data)
        {
            if(resultCode)
                return;
            this["initialize"](data);
            this["displayName"] = "mailbox";
            window.setTimeout(this.$$d__callAppReadyCallback$p$0,0)
        },
        _callAppReadyCallback$p$0: function()
        {
            this._appReadyCallback$p$0()
        },
        _invokeGetTokenMethodAsync$p$0: function(outlookDispid, data, methodName, callback, userContext)
        {
            if($h.ScriptHelpers.isNullOrUndefined(callback))
                throw Error.argumentNull("callback");
            var $$t_9 = this;
            this.invokeHostMethod(outlookDispid,data,function(resultCode, response)
            {
                var asyncResult;
                if(resultCode)
                    asyncResult = $$t_9.createAsyncResult(null,1,9017,userContext,String.format(window["_u"]["ExtensibilityStrings"]["l_InternalProtocolError_Text"],resultCode));
                else
                {
                    var responseDictionary = response;
                    if(window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.get__appName$i$0() === 8 && "error" in responseDictionary && "errorCode" in responseDictionary && responseDictionary["error"] && responseDictionary["errorCode"] === 9030)
                        asyncResult = $$t_9.createAsyncResult(null,1,responseDictionary["errorCode"],userContext,responseDictionary["errorMessage"]);
                    else if(responseDictionary["wasSuccessful"])
                        asyncResult = $$t_9.createAsyncResult(responseDictionary["token"],0,0,userContext,null);
                    else
                        asyncResult = $$t_9.createAsyncResult(null,1,responseDictionary["errorCode"],userContext,responseDictionary["errorMessage"]);
                    if("diagnostics" in responseDictionary)
                        asyncResult["diagnostics"] = responseDictionary["diagnostics"]
                }
                callback(asyncResult)
            })
        },
        getItem: function()
        {
            return this._item$p$0
        },
        _getUserProfile$p$0: function()
        {
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._initialData$p$0._permissionLevel$p$0);
            return this._userProfile$p$0
        },
        _getDiagnostics$p$0: function()
        {
            return this._diagnostics$p$0
        },
        _getEwsUrl$p$0: function()
        {
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._initialData$p$0._permissionLevel$p$0);
            return this._initialData$p$0.get__ewsUrl$i$0()
        },
        _getRestUrl$p$0: function()
        {
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._initialData$p$0._permissionLevel$p$0);
            if(this._shouldInferRestUrl$p$0())
                return this._inferRestUrlFromEwsUrl$p$0();
            return this._initialData$p$0.get__restUrl$i$0()
        },
        _getMasterCategories$p$0: function()
        {
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._initialData$p$0._permissionLevel$p$0);
            if(!this._masterCategories$p$0)
                this._masterCategories$p$0 = new $h.MasterCategories;
            return this._masterCategories$p$0
        },
        _shouldInferRestUrl$p$0: function()
        {
            return window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.get__appName$i$0() === 8 && !this._initialData$p$0.get__restUrl$i$0() && this.isApiVersionSupported("1.5") && this._isHostBuildNumberLessThan$p$0("16.0.8414.1000")
        },
        _isHostBuildNumberLessThan$p$0: function(buildNumber)
        {
            var hostVersion = this._initialData$p$0.get__hostVersion$i$0();
            if(hostVersion)
            {
                var hostVersionParts = hostVersion.split(".");
                var buildNumberParts = buildNumber.split(".");
                return window["parseInt"](hostVersionParts[0]) < window["parseInt"](buildNumberParts[0]) || window["parseInt"](hostVersionParts[0]) === window["parseInt"](buildNumberParts[0]) && window["parseInt"](hostVersionParts[2]) < window["parseInt"](buildNumberParts[2])
            }
            return false
        },
        _inferRestUrlFromEwsUrl$p$0: function()
        {
            var inferredRestUrl = "";
            var stringToFind = "/ews/";
            var index = this._initialData$p$0.get__ewsUrl$i$0().toLowerCase().indexOf(stringToFind);
            if(index !== -1)
                inferredRestUrl = String.format("{0}/{1}",this._initialData$p$0.get__ewsUrl$i$0().slice(0,index),"api");
            return $h.ScriptHelpers.isNonEmptyString(inferredRestUrl) ? inferredRestUrl : null
        },
        _findOffset$p$0: function(value)
        {
            var ranges = this._initialData$p$0.get__timeZoneOffsets$i$0();
            for(var r = 0; r < ranges["length"]; r++)
            {
                var range = ranges[r];
                var start = window["parseInt"](range["start"]);
                var end = window["parseInt"](range["end"]);
                if(value["getTime"]() - start >= 0 && value["getTime"]() - end < 0)
                    return window["parseInt"](range["offset"])
            }
            throw Error.format(window["_u"]["ExtensibilityStrings"]["l_InvalidDate_Text"]);
        },
        _areStringsLoaded$p$0: function()
        {
            var stringsLoaded = false;
            try
            {
                stringsLoaded = !$h.ScriptHelpers.isNullOrUndefined(window["_u"]["ExtensibilityStrings"]["l_EwsRequestOversized_Text"])
            }
            catch($$e_1){}
            return stringsLoaded
        },
        _loadLocalizedScript$p$0: function(stringLoadedCallback)
        {
            var url = null;
            var baseUrl = "";
            var scripts = document.getElementsByTagName("script");
            for(var i = scripts.length - 1; i >= 0; i--)
            {
                var filename = null;
                var attributes = scripts[i].attributes;
                if(attributes)
                {
                    var attribute = attributes.getNamedItem("src");
                    if(attribute)
                        filename = attribute.value;
                    if(filename)
                    {
                        var debug = false;
                        filename = filename.toLowerCase();
                        var officeIndex = filename.indexOf("office_strings.js");
                        if(officeIndex < 0)
                        {
                            officeIndex = filename.indexOf("office_strings.debug.js");
                            debug = true
                        }
                        if(officeIndex > 0 && officeIndex < filename.length)
                        {
                            url = filename.replace(debug ? "office_strings.debug.js" : "office_strings.js","outlook_strings.js");
                            var languageUrl = filename.substring(0,officeIndex);
                            var lastIndexOfSlash = languageUrl.lastIndexOf("/",languageUrl.length - 2);
                            if(lastIndexOfSlash === -1)
                                lastIndexOfSlash = languageUrl.lastIndexOf("\\",languageUrl.length - 2);
                            if(lastIndexOfSlash !== -1 && languageUrl.length > lastIndexOfSlash + 1)
                                baseUrl = languageUrl.substring(0,lastIndexOfSlash + 1);
                            break
                        }
                    }
                }
            }
            if(url)
            {
                var head = document.getElementsByTagName("head")[0];
                var scriptElement = null;
                var $$t_H = this;
                var scriptElementCallback = function()
                    {
                        if(stringLoadedCallback && (!scriptElement.readyState || scriptElement.readyState && (scriptElement.readyState === "loaded" || scriptElement.readyState === "complete")))
                        {
                            scriptElement.onload = null;
                            scriptElement.onreadystatechange = null;
                            stringLoadedCallback()
                        }
                    };
                var $$t_I = this;
                var failureCallback = function()
                    {
                        if(!$$t_I._areStringsLoaded$p$0())
                        {
                            var fallbackUrl = baseUrl + "en-us/" + "outlook_strings.js";
                            scriptElement.onload = null;
                            scriptElement.onreadystatechange = null;
                            scriptElement = $$t_I._createScriptElement$p$0(fallbackUrl);
                            scriptElement.onload = scriptElementCallback;
                            scriptElement.onreadystatechange = scriptElementCallback;
                            head.appendChild(scriptElement)
                        }
                    };
                scriptElement = this._createScriptElement$p$0(url);
                scriptElement.onload = scriptElementCallback;
                scriptElement.onreadystatechange = scriptElementCallback;
                window.setTimeout(failureCallback,2e3);
                head.appendChild(scriptElement)
            }
        },
        _createScriptElement$p$0: function(url)
        {
            var scriptElement = document.createElement("script");
            scriptElement.type = "text/javascript";
            scriptElement.src = url;
            return scriptElement
        }
    };
    OSF.DDA.OutlookAppOm.prototype.initialize = function(initialData)
    {
        if(!initialData)
        {
            this._additionalOutlookParams$p$0 = new $h.AdditionalGlobalParameters(true);
            this._initialData$p$0 = null;
            this._item$p$0 = null;
            return
        }
        var ItemTypeKey = "itemType";
        this._initialData$p$0 = new $h.InitialData(initialData);
        this._hostItemType$p$0 = initialData[ItemTypeKey];
        if(1 === initialData[ItemTypeKey])
            this._item$p$0 = new $h.Message(this._initialData$p$0);
        else if(3 === initialData[ItemTypeKey])
            this._item$p$0 = new $h.MeetingRequest(this._initialData$p$0);
        else if(2 === initialData[ItemTypeKey])
            this._item$p$0 = new $h.Appointment(this._initialData$p$0);
        else if(4 === initialData[ItemTypeKey])
            this._item$p$0 = new $h.MessageCompose(this._initialData$p$0);
        else if(5 === initialData[ItemTypeKey])
            this._item$p$0 = new $h.AppointmentCompose(this._initialData$p$0);
        else if(6 === initialData[ItemTypeKey]);
        else
            Sys.Debug.trace("Unexpected item type was received from the host.");
        this._userProfile$p$0 = new $h.UserProfile(this._initialData$p$0);
        this._diagnostics$p$0 = new $h.Diagnostics(this._initialData$p$0,this._officeAppContext$p$0["get_appName"]());
        var supportsAdditionalParameters = window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.get__appName$i$0() !== 8 || this.isOutlook16OrGreater() || this.isApiVersionSupported("1.5");
        this._additionalOutlookParams$p$0 = new $h.AdditionalGlobalParameters(supportsAdditionalParameters);
        if("itemNumber" in initialData)
            this["setCurrentItemNumber"](initialData["itemNumber"]);
        this._initializeMethods$p$0();
        $h.InitialData._defineReadOnlyProperty$i(this,"item",this.$$d_getItem);
        $h.InitialData._defineReadOnlyProperty$i(this,"userProfile",this.$$d__getUserProfile$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"diagnostics",this.$$d__getDiagnostics$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"ewsUrl",this.$$d__getEwsUrl$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"restUrl",this.$$d__getRestUrl$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"masterCategories",this.$$d__getMasterCategories$p$0);
        if(window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.get__appName$i$0() === 64)
            if(this._initialData$p$0.get__overrideWindowOpen$i$0())
                window.open = this.$$d_windowOpenOverrideHandler;
        this.addEventSupport()
    };
    OSF.DDA.OutlookAppOm.prototype.makeEwsRequestAsync = function(data)
    {
        var args = [];
        for(var $$pai_5 = 1; $$pai_5 < arguments["length"]; ++$$pai_5)
            args[$$pai_5 - 1] = arguments[$$pai_5];
        if($h.ScriptHelpers.isNullOrUndefined(data))
            throw Error.argumentNull("data");
        if(data.length > window["OSF"]["DDA"]["OutlookAppOm"]._maxEwsRequestSize$p)
            throw Error.argument("data",window["_u"]["ExtensibilityStrings"]["l_EwsRequestOversized_Text"]);
        this._throwOnMethodCallForInsufficientPermission$i$0(3,"makeEwsRequestAsync");
        var parameters = $h.CommonParameters.parse(args,true,true);
        var ewsRequest = new $h.EwsRequest(parameters._asyncContext$p$0);
        var $$t_4 = this;
        ewsRequest.onreadystatechange = function()
        {
            if(4 === ewsRequest.get__requestState$i$1())
                parameters._callback$p$0(ewsRequest._asyncResult$p$0)
        };
        ewsRequest.send(data)
    };
    OSF.DDA.OutlookAppOm.prototype.recordDataPoint = function(data)
    {
        if($h.ScriptHelpers.isNullOrUndefined(data))
            throw Error.argumentNull("data");
        this.invokeHostMethod(402,data,null)
    };
    OSF.DDA.OutlookAppOm.prototype.recordTrace = function(data)
    {
        if($h.ScriptHelpers.isNullOrUndefined(data))
            throw Error.argumentNull("data");
        this.invokeHostMethod(401,data,null)
    };
    OSF.DDA.OutlookAppOm.prototype.trackCtq = function(data)
    {
        if($h.ScriptHelpers.isNullOrUndefined(data))
            throw Error.argumentNull("data");
        this.invokeHostMethod(400,data,null)
    };
    OSF.DDA.OutlookAppOm.prototype.convertToLocalClientTime = function(timeValue)
    {
        var date = new Date(timeValue["getTime"]());
        var offset = date["getTimezoneOffset"]() * -1;
        if(this._initialData$p$0 && this._initialData$p$0.get__timeZoneOffsets$i$0())
        {
            date["setUTCMinutes"](date["getUTCMinutes"]() - offset);
            offset = this._findOffset$p$0(date);
            date["setUTCMinutes"](date["getUTCMinutes"]() + offset)
        }
        var retValue = this._dateToDictionary$i$0(date);
        retValue["timezoneOffset"] = offset;
        return retValue
    };
    OSF.DDA.OutlookAppOm.prototype.convertToUtcClientTime = function(input)
    {
        var retValue = this._dictionaryToDate$i$0(input);
        if(this._initialData$p$0 && this._initialData$p$0.get__timeZoneOffsets$i$0())
        {
            var offset = this._findOffset$p$0(retValue);
            retValue["setUTCMinutes"](retValue["getUTCMinutes"]() - offset);
            offset = !input["timezoneOffset"] ? retValue["getTimezoneOffset"]() * -1 : input["timezoneOffset"];
            retValue["setUTCMinutes"](retValue["getUTCMinutes"]() + offset)
        }
        return retValue
    };
    OSF.DDA.OutlookAppOm.prototype.convertToRestId = function(itemId, restVersion)
    {
        if(!itemId)
            throw Error.argumentNull("itemId");
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidRestVersion$p(restVersion);
        return itemId.replace(new RegExp("[/]","g"),"-").replace(new RegExp("[+]","g"),"_")
    };
    OSF.DDA.OutlookAppOm.prototype.convertToEwsId = function(itemId, restVersion)
    {
        if(!itemId)
            throw Error.argumentNull("itemId");
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidRestVersion$p(restVersion);
        return itemId.replace(new RegExp("[-]","g"),"/").replace(new RegExp("[_]","g"),"+")
    };
    OSF.DDA.OutlookAppOm.prototype.getUserIdentityTokenAsync = function()
    {
        var args = [];
        for(var $$pai_2 = 0; $$pai_2 < arguments["length"]; ++$$pai_2)
            args[$$pai_2] = arguments[$$pai_2];
        this._throwOnMethodCallForInsufficientPermission$i$0(1,"getUserIdentityTokenAsync");
        var parameters = $h.CommonParameters.parse(args,true,true);
        this._invokeGetTokenMethodAsync$p$0(2,null,"GetUserIdentityToken",parameters._callback$p$0,parameters._asyncContext$p$0)
    };
    OSF.DDA.OutlookAppOm.prototype.getCallbackTokenAsync = function()
    {
        var args = [];
        for(var $$pai_7 = 0; $$pai_7 < arguments["length"]; ++$$pai_7)
            args[$$pai_7] = arguments[$$pai_7];
        this._throwOnMethodCallForInsufficientPermission$i$0(1,"getCallbackTokenAsync");
        var parameters = $h.CommonParameters.parse(args,true,true);
        var options = {};
        if(parameters._options$p$0)
            for(var $$arr_3 = Object["keys"](parameters._options$p$0), $$len_4 = $$arr_3.length, $$idx_5 = 0; $$idx_5 < $$len_4; ++$$idx_5)
            {
                var key = $$arr_3[$$idx_5];
                options[key] = parameters._options$p$0[key]
            }
        if(!("isRest" in options))
            options["isRest"] = false;
        this._invokeGetTokenMethodAsync$p$0(12,options,"GetCallbackToken",parameters._callback$p$0,parameters._asyncContext$p$0)
    };
    OSF.DDA.OutlookAppOm.prototype.displayMessageForm = function(itemId)
    {
        if($h.ScriptHelpers.isNullOrUndefined(itemId))
            throw Error.argumentNull("itemId");
        this.invokeHostMethod(8,{itemId: window["OSF"]["DDA"]["OutlookAppOm"].getItemIdBasedOnHost(itemId)},null)
    };
    OSF.DDA.OutlookAppOm.prototype.displayAppointmentForm = function(itemId)
    {
        if($h.ScriptHelpers.isNullOrUndefined(itemId))
            throw Error.argumentNull("itemId");
        this.invokeHostMethod(9,{itemId: window["OSF"]["DDA"]["OutlookAppOm"].getItemIdBasedOnHost(itemId)},null)
    };
    OSF.DDA.OutlookAppOm.prototype.logTelemetry = function(jsonData)
    {
        if($h.ScriptHelpers.isNullOrUndefined(jsonData))
            throw Error.argumentNull("jsonData");
        this.invokeHostMethod(163,{telemetryData: jsonData},null)
    };
    OSF.DDA.OutlookAppOm.prototype.RegisterConsentAsync = function(consentState)
    {
        if(consentState !== 2 && consentState !== 1 && consentState)
            throw Error.argumentOutOfRange("consentState");
        var parameters = {};
        parameters["consentState"] = consentState["toString"]();
        parameters["extensionId"] = this["GetExtensionId"]();
        this.invokeHostMethod(40,parameters,null)
    };
    OSF.DDA.OutlookAppOm.prototype.CloseApp = function()
    {
        this.invokeHostMethod(42,null,null)
    };
    OSF.DDA.OutlookAppOm.prototype.GetIsRead = function()
    {
        return this._initialData$p$0.get__isRead$i$0()
    };
    OSF.DDA.OutlookAppOm.prototype.GetEndNodeUrl = function()
    {
        return this._initialData$p$0.get__endNodeUrl$i$0()
    };
    OSF.DDA.OutlookAppOm.prototype.GetConsentMetadata = function()
    {
        return this._initialData$p$0.get__consentMetadata$i$0()
    };
    OSF.DDA.OutlookAppOm.prototype.GetEntryPointUrl = function()
    {
        return this._initialData$p$0.get__entryPointUrl$i$0()
    };
    OSF.DDA.OutlookAppOm.prototype.GetMarketplaceContentMarket = function()
    {
        return this._initialData$p$0.get__marketplaceContentMarket$i$0()
    };
    OSF.DDA.OutlookAppOm.prototype.GetMarketplaceAssetId = function()
    {
        return this._initialData$p$0.get__marketplaceAssetId$i$0()
    };
    OSF.DDA.OutlookAppOm.prototype.GetExtensionId = function()
    {
        return this._initialData$p$0.get__extensionId$i$0()
    };
    OSF.DDA.OutlookAppOm.prototype.setCurrentItemNumber = function(itemNumber)
    {
        this._additionalOutlookParams$p$0.setCurrentItemNumber(itemNumber)
    };
    window["OSF"]["DDA"]["Settings"] = OSF.DDA.Settings = function(data)
    {
        this._rawData$p$0 = data
    };
    window["OSF"]["DDA"]["Settings"]._convertFromRawSettings$p = function(rawSettings)
    {
        if(!rawSettings)
            return{};
        if(window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.get__appName$i$0() === 8 || window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.get__appName$i$0() === 65536 || window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.get__appName$i$0() === 4194304)
        {
            var outlookSettings = rawSettings["SettingsKey"];
            if(outlookSettings)
                return OSF.DDA.SettingsManager["deserializeSettings"](outlookSettings)
        }
        return rawSettings
    };
    OSF.DDA.Settings.prototype = {
        _rawData$p$0: null,
        _settingsData$p$0: null,
        get__data$p$0: function()
        {
            if(!this._settingsData$p$0)
            {
                this._settingsData$p$0 = window["OSF"]["DDA"]["Settings"]._convertFromRawSettings$p(this._rawData$p$0);
                this._rawData$p$0 = null
            }
            return this._settingsData$p$0
        },
        _saveSettingsForOutlook$p$0: function(callback, userContext)
        {
            var storedException = null;
            var startTime = (new Date)["getTime"]();
            var detailedErrorCode = -1;
            try
            {
                var serializedSettings = OSF.DDA.SettingsManager["serializeSettings"](this.get__data$p$0());
                var jsonSettings = window["JSON"]["stringify"](serializedSettings);
                var settingsObjectToSave = {SettingsKey: jsonSettings};
                OSF.DDA.ClientSettingsManager["write"](settingsObjectToSave)
            }
            catch(ex)
            {
                storedException = ex
            }
            var asyncResult;
            if(storedException)
            {
                detailedErrorCode = 9019;
                asyncResult = window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.createAsyncResult(null,1,detailedErrorCode,userContext,storedException["message"])
            }
            else
            {
                detailedErrorCode = 0;
                asyncResult = window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.createAsyncResult(null,0,detailedErrorCode,userContext,null)
            }
            window["OSF"]["AppTelemetry"]["onMethodDone"](404,null,Math["abs"]((new Date)["getTime"]() - startTime),detailedErrorCode);
            if(callback)
                callback(asyncResult)
        },
        _saveSettingsForOwa$p$0: function(callback, userContext)
        {
            var serializedSettings = OSF.DDA.SettingsManager["serializeSettings"](this.get__data$p$0());
            var $$t_7 = this;
            window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.invokeHostMethod(404,[serializedSettings],function(resultCode, response)
            {
                if(callback)
                {
                    var asyncResult;
                    if(resultCode)
                        asyncResult = window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.createAsyncResult(null,1,9017,userContext,String.format(window["_u"]["ExtensibilityStrings"]["l_InternalProtocolError_Text"],resultCode));
                    else
                    {
                        var responseDictionary = response;
                        if(!responseDictionary["error"])
                            asyncResult = window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.createAsyncResult(null,0,0,userContext,null);
                        else
                            asyncResult = window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.createAsyncResult(null,1,9019,userContext,responseDictionary["errorMessage"])
                    }
                    callback(asyncResult)
                }
            })
        }
    };
    OSF.DDA.Settings.prototype.get = function(name)
    {
        return this.get__data$p$0()[name]
    };
    OSF.DDA.Settings.prototype.set = function(name, value)
    {
        this.get__data$p$0()[name] = value
    };
    OSF.DDA.Settings.prototype.remove = function(name)
    {
        delete this.get__data$p$0()[name]
    };
    OSF.DDA.Settings.prototype.saveAsync = function()
    {
        var args = [];
        for(var $$pai_4 = 0; $$pai_4 < arguments["length"]; ++$$pai_4)
            args[$$pai_4] = arguments[$$pai_4];
        var commonParameters = $h.CommonParameters.parse(args,false);
        if(window["JSON"]["stringify"](OSF.DDA.SettingsManager["serializeSettings"](this.get__data$p$0())).length > 32768)
        {
            var asyncResult = window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.createAsyncResult(null,1,9019,commonParameters._asyncContext$p$0,"");
            var $$t_3 = this;
            window.setTimeout(function()
            {
                commonParameters._callback$p$0(asyncResult)
            },0);
            return
        }
        if(window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.get__appName$i$0() === 64)
            this._saveSettingsForOwa$p$0(commonParameters._callback$p$0,commonParameters._asyncContext$p$0);
        else
            this._saveSettingsForOutlook$p$0(commonParameters._callback$p$0,commonParameters._asyncContext$p$0)
    };
    Type.registerNamespace("$h");
    var $h = window["$h"] || {};
    Type.registerNamespace("Office.cast");
    var Office = window["Office"] || {};
    Office.cast = Office.cast || {};
    $h.AdditionalGlobalParameters = function(supported)
    {
        this._parameterBlobSupported$p$0 = supported;
        this._itemNumber$p$0 = 0
    };
    $h.AdditionalGlobalParameters.prototype = {
        _parameterBlobSupported$p$0: false,
        _itemNumber$p$0: 0,
        _actionsDefinition$p$0: null,
        setActionsDefinition: function(actionsDefinition)
        {
            this._actionsDefinition$p$0 = actionsDefinition
        },
        setCurrentItemNumber: function(itemNumber)
        {
            if(itemNumber > 0)
                this._itemNumber$p$0 = itemNumber
        },
        get_itemNumber: function()
        {
            return this._itemNumber$p$0
        },
        get_actionsDefinition: function()
        {
            return this._actionsDefinition$p$0
        },
        updateOutlookExecuteParameters: function(executeParameters, additionalApiParameters)
        {
            if(this._parameterBlobSupported$p$0)
            {
                if(this._itemNumber$p$0 > 0)
                    additionalApiParameters["itemNumber"] = this._itemNumber$p$0["toString"]();
                if(this._actionsDefinition$p$0)
                    additionalApiParameters["actions"] = this._actionsDefinition$p$0;
                if(!Object["keys"](additionalApiParameters)["length"])
                    return;
                if(!executeParameters["val"])
                    executeParameters["val"] = [];
                executeParameters["val"]["push"](window["JSON"]["stringify"](additionalApiParameters))
            }
        }
    };
    $h.Appointment = function(dataDictionary)
    {
        this.$$d__getEnhancedLocation$p$2 = Function.createDelegate(this,this._getEnhancedLocation$p$2);
        this.$$d__getSeriesId$p$2 = Function.createDelegate(this,this._getSeriesId$p$2);
        this.$$d__getRecurrence$p$2 = Function.createDelegate(this,this._getRecurrence$p$2);
        this.$$d__getOrganizer$p$2 = Function.createDelegate(this,this._getOrganizer$p$2);
        this.$$d__getNormalizedSubject$p$2 = Function.createDelegate(this,this._getNormalizedSubject$p$2);
        this.$$d__getSubject$p$2 = Function.createDelegate(this,this._getSubject$p$2);
        this.$$d__getResources$p$2 = Function.createDelegate(this,this._getResources$p$2);
        this.$$d__getRequiredAttendees$p$2 = Function.createDelegate(this,this._getRequiredAttendees$p$2);
        this.$$d__getOptionalAttendees$p$2 = Function.createDelegate(this,this._getOptionalAttendees$p$2);
        this.$$d__getLocation$p$2 = Function.createDelegate(this,this._getLocation$p$2);
        this.$$d__getEnd$p$2 = Function.createDelegate(this,this._getEnd$p$2);
        this.$$d__getStart$p$2 = Function.createDelegate(this,this._getStart$p$2);
        $h.Appointment["initializeBase"](this,[dataDictionary]);
        $h.InitialData._defineReadOnlyProperty$i(this,"start",this.$$d__getStart$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"end",this.$$d__getEnd$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"location",this.$$d__getLocation$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"optionalAttendees",this.$$d__getOptionalAttendees$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"requiredAttendees",this.$$d__getRequiredAttendees$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"resources",this.$$d__getResources$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"subject",this.$$d__getSubject$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"normalizedSubject",this.$$d__getNormalizedSubject$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"organizer",this.$$d__getOrganizer$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"recurrence",this.$$d__getRecurrence$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"seriesId",this.$$d__getSeriesId$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"enhancedLocation",this.$$d__getEnhancedLocation$p$2)
    };
    $h.Appointment.prototype = {
        _enhancedLocation$p$2: null,
        getItemType: function()
        {
            return window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ItemType"]["Appointment"]
        },
        _getStart$p$2: function()
        {
            return this._data$p$0.get__start$i$0()
        },
        _getEnd$p$2: function()
        {
            return this._data$p$0.get__end$i$0()
        },
        _getLocation$p$2: function()
        {
            return this._data$p$0.get__location$i$0()
        },
        _getOptionalAttendees$p$2: function()
        {
            return this._data$p$0.get__cc$i$0()
        },
        _getRequiredAttendees$p$2: function()
        {
            return this._data$p$0.get__to$i$0()
        },
        _getResources$p$2: function()
        {
            return this._data$p$0.get__resources$i$0()
        },
        _getSubject$p$2: function()
        {
            return this._data$p$0.get__subject$i$0()
        },
        _getNormalizedSubject$p$2: function()
        {
            return this._data$p$0.get__normalizedSubject$i$0()
        },
        _getOrganizer$p$2: function()
        {
            return this._data$p$0.get__organizer$i$0()
        },
        _getRecurrence$p$2: function()
        {
            if(this._data$p$0.get__recurrence$i$0() && this._data$p$0.get__recurrence$i$0()["seriesTimeJson"])
                return $h.ComposeRecurrence.copyRecurrenceObjectConvertSeriesTimeJson(this._data$p$0.get__recurrence$i$0());
            return this._data$p$0.get__recurrence$i$0()
        },
        _getSeriesId$p$2: function()
        {
            return this._data$p$0.get__seriesId$i$0()
        },
        _getEnhancedLocation$p$2: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._enhancedLocation$p$2)
                this._enhancedLocation$p$2 = new $h.EnhancedLocation(false);
            return this._enhancedLocation$p$2
        }
    };
    $h.Appointment.prototype.getEntities = function()
    {
        return this._data$p$0._getEntities$i$0()
    };
    $h.Appointment.prototype.getEntitiesByType = function(entityType)
    {
        return this._data$p$0._getEntitiesByType$i$0(entityType)
    };
    $h.Appointment.prototype.getSelectedEntities = function()
    {
        return this._data$p$0._getSelectedEntities$i$0()
    };
    $h.Appointment.prototype.getRegExMatches = function()
    {
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"getRegExMatches");
        return this._data$p$0._getRegExMatches$i$0()
    };
    $h.Appointment.prototype.getFilteredEntitiesByName = function(name)
    {
        return this._data$p$0._getFilteredEntitiesByName$i$0(name)
    };
    $h.Appointment.prototype.getRegExMatchesByName = function(name)
    {
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"getRegExMatchesByName");
        return this._data$p$0._getRegExMatchesByName$i$0(name)
    };
    $h.Appointment.prototype.getSelectedRegExMatches = function()
    {
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"getSelectedRegExMatches");
        return this._data$p$0._getSelectedRegExMatches$i$0()
    };
    $h.Appointment.prototype.displayReplyForm = function(obj)
    {
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._displayReplyForm$i$0(obj)
    };
    $h.Appointment.prototype.displayReplyAllForm = function(obj)
    {
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._displayReplyAllForm$i$0(obj)
    };
    $h.AppointmentCompose = function(data)
    {
        this.$$d__getEnhancedLocation$p$2 = Function.createDelegate(this,this._getEnhancedLocation$p$2);
        this.$$d__getOrganizer$p$2 = Function.createDelegate(this,this._getOrganizer$p$2);
        this.$$d__getSeriesId$p$2 = Function.createDelegate(this,this._getSeriesId$p$2);
        this.$$d__getRecurrence$p$2 = Function.createDelegate(this,this._getRecurrence$p$2);
        this.$$d__getLocation$p$2 = Function.createDelegate(this,this._getLocation$p$2);
        this.$$d__getEnd$p$2 = Function.createDelegate(this,this._getEnd$p$2);
        this.$$d__getStart$p$2 = Function.createDelegate(this,this._getStart$p$2);
        this.$$d__getOptionalAttendees$p$2 = Function.createDelegate(this,this._getOptionalAttendees$p$2);
        this.$$d__getRequiredAttendees$p$2 = Function.createDelegate(this,this._getRequiredAttendees$p$2);
        $h.AppointmentCompose["initializeBase"](this,[data]);
        $h.InitialData._defineReadOnlyProperty$i(this,"requiredAttendees",this.$$d__getRequiredAttendees$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"optionalAttendees",this.$$d__getOptionalAttendees$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"start",this.$$d__getStart$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"end",this.$$d__getEnd$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"location",this.$$d__getLocation$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"recurrence",this.$$d__getRecurrence$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"seriesId",this.$$d__getSeriesId$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"organizer",this.$$d__getOrganizer$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"enhancedLocation",this.$$d__getEnhancedLocation$p$2)
    };
    $h.AppointmentCompose.prototype = {
        _requiredAttendees$p$2: null,
        _optionalAttendees$p$2: null,
        _start$p$2: null,
        _end$p$2: null,
        _location$p$2: null,
        _enhancedLocation$p$2: null,
        _recurrence$p$2: null,
        _organizer$p$2: null,
        getItemType: function()
        {
            return window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ItemType"]["Appointment"]
        },
        _getRequiredAttendees$p$2: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._requiredAttendees$p$2)
                this._requiredAttendees$p$2 = new $h.ComposeRecipient(0,"requiredAttendees");
            return this._requiredAttendees$p$2
        },
        _getOptionalAttendees$p$2: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._optionalAttendees$p$2)
                this._optionalAttendees$p$2 = new $h.ComposeRecipient(1,"optionalAttendees");
            return this._optionalAttendees$p$2
        },
        _getStart$p$2: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._start$p$2)
                this._start$p$2 = new $h.ComposeTime(1);
            return this._start$p$2
        },
        _getEnd$p$2: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._end$p$2)
                this._end$p$2 = new $h.ComposeTime(2);
            return this._end$p$2
        },
        _getLocation$p$2: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._location$p$2)
                this._location$p$2 = new $h.ComposeLocation;
            return this._location$p$2
        },
        _getEnhancedLocation$p$2: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._enhancedLocation$p$2)
                this._enhancedLocation$p$2 = new $h.EnhancedLocation(true);
            return this._enhancedLocation$p$2
        },
        _getRecurrence$p$2: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._recurrence$p$2)
            {
                var isInstance = !!this._data$p$0.get__seriesId$i$0() && this._data$p$0.get__seriesId$i$0().length > 0;
                this._recurrence$p$2 = new $h.ComposeRecurrence(isInstance)
            }
            return this._recurrence$p$2
        },
        _getSeriesId$p$2: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            return this._data$p$0.get__seriesId$i$0()
        },
        _getOrganizer$p$2: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._organizer$p$2)
                this._organizer$p$2 = new $h.ComposeFrom;
            return this._organizer$p$2
        }
    };
    $h.AttachmentConstants = function(){};
    $h.AttachmentDetails = function(data)
    {
        this.$$d__getUrl$p$0 = Function.createDelegate(this,this._getUrl$p$0);
        this.$$d__getIsInline$p$0 = Function.createDelegate(this,this._getIsInline$p$0);
        this.$$d__getAttachmentType$p$0 = Function.createDelegate(this,this._getAttachmentType$p$0);
        this.$$d__getSize$p$0 = Function.createDelegate(this,this._getSize$p$0);
        this.$$d__getContentType$p$0 = Function.createDelegate(this,this._getContentType$p$0);
        this.$$d__getName$p$0 = Function.createDelegate(this,this._getName$p$0);
        this.$$d__getId$p$0 = Function.createDelegate(this,this._getId$p$0);
        this._data$p$0 = data;
        $h.InitialData._defineReadOnlyProperty$i(this,"id",this.$$d__getId$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"name",this.$$d__getName$p$0);
        if("contentType" in this._data$p$0)
            $h.InitialData._defineReadOnlyProperty$i(this,"contentType",this.$$d__getContentType$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"size",this.$$d__getSize$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"attachmentType",this.$$d__getAttachmentType$p$0);
        if("isInline" in this._data$p$0)
            $h.InitialData._defineReadOnlyProperty$i(this,"isInline",this.$$d__getIsInline$p$0);
        if("url" in this._data$p$0)
            $h.InitialData._defineReadOnlyProperty$i(this,"url",this.$$d__getUrl$p$0)
    };
    $h.AttachmentDetails.createFromJsonArray = function(arrayOfAttachmentJsonData)
    {
        var attachmentJsonArray = arrayOfAttachmentJsonData;
        if($h.ScriptHelpers.isNullOrUndefined(attachmentJsonArray))
            return new Array(0);
        var attachmentDetails = [];
        for(var i = 0; i < attachmentJsonArray["length"]; i++)
            if(!$h.ScriptHelpers.isNullOrUndefined(attachmentJsonArray[i]))
                attachmentDetails["push"](new $h.AttachmentDetails(attachmentJsonArray[i]));
        return attachmentDetails
    };
    $h.AttachmentDetails.prototype = {
        _data$p$0: null,
        _getId$p$0: function()
        {
            return this._data$p$0["id"]
        },
        _getName$p$0: function()
        {
            return this._data$p$0["name"]
        },
        _getContentType$p$0: function()
        {
            return this._data$p$0["contentType"]
        },
        _getSize$p$0: function()
        {
            return this._data$p$0["size"]
        },
        _getAttachmentType$p$0: function()
        {
            var response = this._data$p$0["attachmentType"];
            return response < $h.AttachmentDetails._attachmentTypeMap$p["length"] ? $h.AttachmentDetails._attachmentTypeMap$p[response] : window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["AttachmentType"]["File"]
        },
        _getIsInline$p$0: function()
        {
            return this._data$p$0["isInline"]
        },
        _getUrl$p$0: function()
        {
            return this._data$p$0["url"]
        }
    };
    $h.Body = function(){};
    $h.Body._tryMapToHostCoercionType$i = function(coercionType, hostCoercionType)
    {
        hostCoercionType["val"] = undefined;
        if(coercionType === window["Microsoft"]["Office"]["WebExtension"]["CoercionType"]["Html"])
            hostCoercionType["val"] = 3;
        else if(coercionType === window["Microsoft"]["Office"]["WebExtension"]["CoercionType"]["Text"])
            hostCoercionType["val"] = 0;
        else
            return false;
        return true
    };
    $h.Body.prototype.getAsync = function(coercionType)
    {
        var args = [];
        for(var $$pai_7 = 1; $$pai_7 < arguments["length"]; ++$$pai_7)
            args[$$pai_7 - 1] = arguments[$$pai_7];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"body.getAsync");
        var commonParameters = $h.CommonParameters.parse(args,true);
        var hostCoercionType;
        var $$t_5,
            $$t_6;
        if(!($$t_6 = $h.Body._tryMapToHostCoercionType$i(coercionType,$$t_5 = {val: hostCoercionType}),hostCoercionType = $$t_5["val"],$$t_6))
            throw Error.argument("coercionType");
        var dataToHost = {coercionType: hostCoercionType};
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(37,dataToHost,null,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
    };
    $h.Categories = function(){};
    $h.Categories.prototype.getAsync = function()
    {
        var args = [];
        for(var $$pai_2 = 0; $$pai_2 < arguments["length"]; ++$$pai_2)
            args[$$pai_2] = arguments[$$pai_2];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"categories.getAsync");
        var parameters = $h.CommonParameters.parse(args,true);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(157,null,null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.Categories.prototype.addAsync = function(categories)
    {
        var args = [];
        for(var $$pai_3 = 1; $$pai_3 < arguments["length"]; ++$$pai_3)
            args[$$pai_3 - 1] = arguments[$$pai_3];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"categories.addAsync");
        var parameters = $h.CommonParameters.parse(args,false);
        $h.ScriptHelpers.validateCategoriesArray(categories);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(158,{categories: categories},null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.Categories.prototype.removeAsync = function(categories)
    {
        var args = [];
        for(var $$pai_3 = 1; $$pai_3 < arguments["length"]; ++$$pai_3)
            args[$$pai_3 - 1] = arguments[$$pai_3];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"categories.removeAsync");
        var parameters = $h.CommonParameters.parse(args,false);
        $h.ScriptHelpers.validateCategoriesArray(categories);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(159,{categories: categories},null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ComposeFrom = function()
    {
        this.$$d__getAsyncFormatter$p$0 = Function.createDelegate(this,this._getAsyncFormatter$p$0)
    };
    $h.ComposeFrom.prototype = {_getAsyncFormatter$p$0: function(rawInput)
        {
            var from = rawInput;
            return $h.ScriptHelpers.isNullOrUndefined(from) ? null : new $h.EmailAddressDetails(from)
        }};
    $h.ComposeFrom.prototype.getAsync = function()
    {
        var args = [];
        for(var $$pai_2 = 0; $$pai_2 < arguments["length"]; ++$$pai_2)
            args[$$pai_2] = arguments[$$pai_2];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"from.getAsync");
        var parameters = $h.CommonParameters.parse(args,true);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(107,null,this.$$d__getAsyncFormatter$p$0,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.InternetHeaders = function(supportsWriteMethods)
    {
        this.$$d_getAsyncApi = Function.createDelegate(this,this.getAsyncApi);
        this.$$d_removeAsyncApi = Function.createDelegate(this,this.removeAsyncApi);
        this.$$d_setAsyncApi = Function.createDelegate(this,this.setAsyncApi);
        var currentInstance = this;
        if(supportsWriteMethods)
        {
            currentInstance["setAsync"] = this.$$d_setAsyncApi;
            currentInstance["removeAsync"] = this.$$d_removeAsyncApi
        }
        currentInstance["getAsync"] = this.$$d_getAsyncApi
    };
    $h.InternetHeaders.prototype = {
        getAsyncApi: function(internetHeadersNames)
        {
            var args = [];
            for(var $$pai_3 = 1; $$pai_3 < arguments["length"]; ++$$pai_3)
                args[$$pai_3 - 1] = arguments[$$pai_3];
            window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"internetHeaders.getAsync");
            var parameters = $h.CommonParameters.parse(args,true);
            this.validateInternetHeaderArray(internetHeadersNames);
            window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(151,{internetHeaderKeys: internetHeadersNames},null,parameters._asyncContext$p$0,parameters._callback$p$0)
        },
        setAsyncApi: function(internetHeadersNameValuePairs)
        {
            var args = [];
            for(var $$pai_7 = 1; $$pai_7 < arguments["length"]; ++$$pai_7)
                args[$$pai_7 - 1] = arguments[$$pai_7];
            window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"internetHeaders.setAsync");
            var parameters = $h.CommonParameters.parse(args,false);
            if($h.ScriptHelpers.isNullOrUndefined(internetHeadersNameValuePairs))
                throw Error.argument("internetHeaders");
            var keys = Object["keys"](internetHeadersNameValuePairs);
            if(!keys["length"])
                throw Error.argument("internetHeaders");
            for(var i = 0; i < keys["length"]; i++)
            {
                var key = keys[i];
                if(!String["isInstanceOfType"](internetHeadersNameValuePairs[key]))
                    throw Error.argument("internetHeaders");
                var value = internetHeadersNameValuePairs[key];
                window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(key.length + value.length,0,998,key)
            }
            window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(152,{internetHeaderNameValuePairs: internetHeadersNameValuePairs},null,parameters._asyncContext$p$0,parameters._callback$p$0)
        },
        removeAsyncApi: function(internetHeadersNames)
        {
            var args = [];
            for(var $$pai_3 = 1; $$pai_3 < arguments["length"]; ++$$pai_3)
                args[$$pai_3 - 1] = arguments[$$pai_3];
            window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"internetHeaders.removeAsync");
            var parameters = $h.CommonParameters.parse(args,false);
            this.validateInternetHeaderArray(internetHeadersNames);
            window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(153,{internetHeaderKeys: internetHeadersNames},null,parameters._asyncContext$p$0,parameters._callback$p$0)
        },
        validateInternetHeaderArray: function(internetHeaderArray)
        {
            if($h.ScriptHelpers.isNullOrUndefined(internetHeaderArray))
                throw Error.argument("internetHeaders");
            if(!Array["isInstanceOfType"](internetHeaderArray))
                throw Error.argumentType("internetHeaders",Object["getType"](internetHeaderArray),Array);
            if(!internetHeaderArray["length"])
                throw Error.argument("internetHeaders");
            for(var i = 0; i < internetHeaderArray["length"]; i++)
                if(!$h.ScriptHelpers.isNonEmptyString(internetHeaderArray[i]))
                    throw Error.argument("internetHeaders");
        }
    };
    $h.ComposeBody = function()
    {
        $h.ComposeBody["initializeBase"](this)
    };
    $h.ComposeBody._createParameterDictionaryToHost$i = function(data, parameters)
    {
        var dataToHost = {data: data};
        return $h.ComposeBody._addCoercionTypeToDictionary$i(dataToHost,parameters)
    };
    $h.ComposeBody._createAppendParameterDictionaryToHost$i = function(data, parameters)
    {
        var dataToHost = {appendTxt: data};
        return $h.ComposeBody._addCoercionTypeToDictionary$i(dataToHost,parameters)
    };
    $h.ComposeBody._addCoercionTypeToDictionary$i = function(dataToHost, parameters)
    {
        if(parameters._options$p$0 && parameters._options$p$0["hasOwnProperty"]("coercionType") && !$h.ScriptHelpers.isNull(parameters._options$p$0["coercionType"]))
        {
            var hostCoercionType;
            var $$t_3,
                $$t_4;
            if(!($$t_4 = $h.Body._tryMapToHostCoercionType$i(parameters._options$p$0["coercionType"],$$t_3 = {val: hostCoercionType}),hostCoercionType = $$t_3["val"],$$t_4))
            {
                if(parameters._callback$p$0)
                    parameters._callback$p$0(window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.createAsyncResult(null,1,1e3,parameters._asyncContext$p$0,null));
                return null
            }
            dataToHost["coercionType"] = hostCoercionType
        }
        else
            dataToHost["coercionType"] = 0;
        return dataToHost
    };
    $h.ComposeBody.prototype.getTypeAsync = function()
    {
        var args = [];
        for(var $$pai_2 = 0; $$pai_2 < arguments["length"]; ++$$pai_2)
            args[$$pai_2] = arguments[$$pai_2];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"body.getTypeAsync");
        var parameters = $h.CommonParameters.parse(args,true);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(14,null,null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ComposeBody.prototype.setSelectedDataAsync = function(data)
    {
        var args = [];
        for(var $$pai_4 = 1; $$pai_4 < arguments["length"]; ++$$pai_4)
            args[$$pai_4 - 1] = arguments[$$pai_4];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"body.setSelectedDataAsync");
        var parameters = $h.CommonParameters.parse(args,false);
        if(!String["isInstanceOfType"](data))
            throw Error.argumentType("data",Object["getType"](data),String);
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(data.length,0,1e6,"data");
        var dataToHost = $h.ComposeBody._createParameterDictionaryToHost$i(data,parameters);
        if(!dataToHost)
            return;
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(13,dataToHost,null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ComposeBody.prototype.prependAsync = function(data)
    {
        var args = [];
        for(var $$pai_4 = 1; $$pai_4 < arguments["length"]; ++$$pai_4)
            args[$$pai_4 - 1] = arguments[$$pai_4];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"body.prependAsync");
        var parameters = $h.CommonParameters.parse(args,false);
        if(!String["isInstanceOfType"](data))
            throw Error.argumentType("data",Object["getType"](data),String);
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(data.length,0,1e6,"data");
        var dataToHost = $h.ComposeBody._createParameterDictionaryToHost$i(data,parameters);
        if(!dataToHost)
            return;
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(23,dataToHost,null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ComposeBody.prototype.appendOnSendAsync = function(data)
    {
        var args = [];
        for(var $$pai_4 = 1; $$pai_4 < arguments["length"]; ++$$pai_4)
            args[$$pai_4 - 1] = arguments[$$pai_4];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"body.appendOnSendAsync");
        var parameters = $h.CommonParameters.parse(args,false);
        if(!data)
            data = "";
        if(!String["isInstanceOfType"](data))
            throw Error.argumentType("data",Object["getType"](data),String);
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(data.length,0,5e3,"data");
        var dataToHost = $h.ComposeBody._createAppendParameterDictionaryToHost$i(data,parameters);
        if(!dataToHost)
            return;
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(100,dataToHost,null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ComposeBody.prototype.setAsync = function(data)
    {
        var args = [];
        for(var $$pai_4 = 1; $$pai_4 < arguments["length"]; ++$$pai_4)
            args[$$pai_4 - 1] = arguments[$$pai_4];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"body.setAsync");
        var parameters = $h.CommonParameters.parse(args,false);
        if(!String["isInstanceOfType"](data))
            throw Error.argumentType("data",Object["getType"](data),String);
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(data.length,0,1e6,"data");
        var dataToHost = $h.ComposeBody._createParameterDictionaryToHost$i(data,parameters);
        if(!dataToHost)
            return;
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(38,dataToHost,null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ComposeItem = function(data)
    {
        this.$$d__getBody$p$1 = Function.createDelegate(this,this._getBody$p$1);
        this.$$d__getSubject$p$1 = Function.createDelegate(this,this._getSubject$p$1);
        $h.ComposeItem["initializeBase"](this,[data]);
        $h.InitialData._defineReadOnlyProperty$i(this,"subject",this.$$d__getSubject$p$1);
        $h.InitialData._defineReadOnlyProperty$i(this,"body",this.$$d__getBody$p$1)
    };
    $h.ComposeItem._validateAndExtractCommonParametersForAddFileAttachmentApis$p = function(attachmentName, args, apiName, parameters, asyncContext, callback)
    {
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,apiName);
        if(!$h.ScriptHelpers.isNonEmptyString(attachmentName))
            throw Error.argument("attachmentName");
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(attachmentName.length,0,255,"attachmentName");
        var commonParameters = $h.CommonParameters.parse(args,false);
        var isInline = false;
        if(!$h.ScriptHelpers.isNull(commonParameters._options$p$0))
            isInline = $h.ScriptHelpers.isValueTrue(commonParameters._options$p$0["isInline"]);
        parameters["val"] = {
            name: attachmentName,
            isInline: isInline,
            __timeout__: 6e5
        };
        asyncContext["val"] = commonParameters._asyncContext$p$0;
        callback["val"] = commonParameters._callback$p$0
    };
    $h.ComposeItem.prototype = {
        _subject$p$1: null,
        _body$p$1: null,
        _getBody$p$1: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._body$p$1)
                this._body$p$1 = new $h.ComposeBody;
            return this._body$p$1
        },
        _getSubject$p$1: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._subject$p$1)
                this._subject$p$1 = new $h.ComposeSubject;
            return this._subject$p$1
        }
    };
    $h.ComposeItem.prototype.addFileAttachmentAsync = function(uri, attachmentName)
    {
        var args = [];
        for(var $$pai_9 = 2; $$pai_9 < arguments["length"]; ++$$pai_9)
            args[$$pai_9 - 2] = arguments[$$pai_9];
        var parameters;
        var asyncContext;
        var callback;
        if(!$h.ScriptHelpers.isNonEmptyString(uri))
            throw Error.argument("uri");
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(uri.length,0,2048,"uri");
        var $$t_6,
            $$t_7,
            $$t_8;
        $h.ComposeItem._validateAndExtractCommonParametersForAddFileAttachmentApis$p(attachmentName,args,"addFileAttachmentAsync",$$t_6 = {val: parameters},$$t_7 = {val: asyncContext},$$t_8 = {val: callback}),parameters = $$t_6["val"],
            asyncContext = $$t_7["val"],
            callback = $$t_8["val"];
        parameters["uri"] = uri;
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(16,parameters,null,asyncContext,callback)
    };
    $h.ComposeItem.prototype.addBase64FileAttachmentAsync = function(base64Encoded, attachmentName)
    {
        var args = [];
        for(var $$pai_9 = 2; $$pai_9 < arguments["length"]; ++$$pai_9)
            args[$$pai_9 - 2] = arguments[$$pai_9];
        var parameters;
        var asyncContext;
        var callback;
        if(!$h.ScriptHelpers.isNonEmptyString(base64Encoded))
            throw Error.argument("base64Encoded");
        var $$t_6,
            $$t_7,
            $$t_8;
        $h.ComposeItem._validateAndExtractCommonParametersForAddFileAttachmentApis$p(attachmentName,args,"addBase64FileAttachmentAsync",$$t_6 = {val: parameters},$$t_7 = {val: asyncContext},$$t_8 = {val: callback}),parameters = $$t_6["val"],
            asyncContext = $$t_7["val"],
            callback = $$t_8["val"];
        parameters["base64String"] = base64Encoded;
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(148,parameters,null,asyncContext,callback)
    };
    $h.ComposeItem.prototype.addItemAttachmentAsync = function(itemId, attachmentName)
    {
        var args = [];
        for(var $$pai_5 = 2; $$pai_5 < arguments["length"]; ++$$pai_5)
            args[$$pai_5 - 2] = arguments[$$pai_5];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"addItemAttachmentAsync");
        if(!$h.ScriptHelpers.isNonEmptyString(itemId))
            throw Error.argument("itemId");
        if(!$h.ScriptHelpers.isNonEmptyString(attachmentName))
            throw Error.argument("attachmentName");
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(itemId.length,0,200,"itemId");
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(attachmentName.length,0,255,"attachmentName");
        var commonParameters = $h.CommonParameters.parse(args,false);
        var parameters = {
                itemId: window["OSF"]["DDA"]["OutlookAppOm"].getItemIdBasedOnHost(itemId),
                name: attachmentName,
                __timeout__: 6e5
            };
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(19,parameters,null,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
    };
    $h.ComposeItem.prototype.getAttachmentsAsync = function()
    {
        var args = [];
        for(var $$pai_2 = 0; $$pai_2 < arguments["length"]; ++$$pai_2)
            args[$$pai_2] = arguments[$$pai_2];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"getAttachmentsAsync");
        var commonParameters = $h.CommonParameters.parse(args,true);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(149,null,$h.AttachmentDetails.createFromJsonArray,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
    };
    $h.ComposeItem.prototype.removeAttachmentAsync = function(attachmentId)
    {
        var args = [];
        for(var $$pai_3 = 1; $$pai_3 < arguments["length"]; ++$$pai_3)
            args[$$pai_3 - 1] = arguments[$$pai_3];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"removeAttachmentAsync");
        if(!$h.ScriptHelpers.isNonEmptyString(attachmentId))
            throw Error.argument("attachmentId");
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(attachmentId.length,0,200,"attachmentId");
        var commonParameters = $h.CommonParameters.parse(args,false);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(20,{attachmentIndex: attachmentId},null,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
    };
    $h.ComposeItem.prototype.getSelectedDataAsync = function(coercionType)
    {
        var args = [];
        for(var $$pai_7 = 1; $$pai_7 < arguments["length"]; ++$$pai_7)
            args[$$pai_7 - 1] = arguments[$$pai_7];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"getSelectedDataAsync");
        var commonParameters = $h.CommonParameters.parse(args,true);
        var hostCoercionType;
        var $$t_5,
            $$t_6;
        if(coercionType !== window["Microsoft"]["Office"]["WebExtension"]["CoercionType"]["Html"] && coercionType !== window["Microsoft"]["Office"]["WebExtension"]["CoercionType"]["Text"] || !($$t_6 = $h.Body._tryMapToHostCoercionType$i(coercionType,$$t_5 = {val: hostCoercionType}),hostCoercionType = $$t_5["val"],$$t_6))
            throw Error.argument("coercionType");
        var dataToHost = {coercionType: hostCoercionType};
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(28,dataToHost,null,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
    };
    $h.ComposeItem.prototype.setSelectedDataAsync = function(data)
    {
        var args = [];
        for(var $$pai_4 = 1; $$pai_4 < arguments["length"]; ++$$pai_4)
            args[$$pai_4 - 1] = arguments[$$pai_4];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"setSelectedDataAsync");
        var parameters = $h.CommonParameters.parse(args,false);
        if(!String["isInstanceOfType"](data))
            throw Error.argumentType("data",Object["getType"](data),String);
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(data.length,0,1e6,"data");
        var dataToHost = $h.ComposeBody._createParameterDictionaryToHost$i(data,parameters);
        if(!dataToHost)
            return;
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(29,dataToHost,null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ComposeItem.prototype.close = function()
    {
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(41,null,null,null,null)
    };
    $h.ComposeItem.prototype.saveAsync = function()
    {
        var args = [];
        for(var $$pai_2 = 0; $$pai_2 < arguments["length"]; ++$$pai_2)
            args[$$pai_2] = arguments[$$pai_2];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"saveAsync");
        var parameters = $h.CommonParameters.parse(args,false);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(32,null,null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ComposeItem.prototype.getItemIdAsync = function()
    {
        var args = [];
        for(var $$pai_2 = 0; $$pai_2 < arguments["length"]; ++$$pai_2)
            args[$$pai_2] = arguments[$$pai_2];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"getItemIdAsync");
        var parameters = $h.CommonParameters.parse(args,false);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(164,null,null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ComposeRecipient = function(type, propertyName)
    {
        this._type$p$0 = type;
        this._propertyName$p$0 = propertyName
    };
    $h.ComposeRecipient._throwOnInvalidDisplayNameOrEmail$p = function(displayName, emailAddress)
    {
        if(!displayName && !emailAddress)
            throw Error.argument("recipients");
        if(displayName && displayName.length > 255)
            throw Error.argumentOutOfRange("recipients",displayName.length,window["_u"]["ExtensibilityStrings"]["l_DisplayNameTooLong_Text"]);
        if(emailAddress && emailAddress.length > 571)
            throw Error.argumentOutOfRange("recipients",emailAddress.length,window["_u"]["ExtensibilityStrings"]["l_EmailAddressTooLong_Text"]);
    };
    $h.ComposeRecipient._getAsyncFormatter$p = function(rawInput)
    {
        var input = rawInput;
        var output = [];
        for(var i = 0; i < input["length"]; i++)
        {
            var email = new $h.EmailAddressDetails(input[i]);
            output[i] = email
        }
        return output
    };
    $h.ComposeRecipient._createEmailDictionaryForHost$p = function(address, name)
    {
        return{
                address: address,
                name: name
            }
    };
    $h.ComposeRecipient.prototype = {
        _propertyName$p$0: null,
        _type$p$0: 0,
        setAddHelper: function(recipients, args, isSet)
        {
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(recipients["length"],0,100,"recipients");
            var parameters = $h.CommonParameters.parse(args,false);
            var recipientData = [];
            if(Array["isInstanceOfType"](recipients))
                for(var i = 0; i < recipients["length"]; i++)
                    if(String["isInstanceOfType"](recipients[i]))
                    {
                        $h.ComposeRecipient._throwOnInvalidDisplayNameOrEmail$p(recipients[i],recipients[i]);
                        recipientData[i] = $h.ComposeRecipient._createEmailDictionaryForHost$p(recipients[i],recipients[i])
                    }
                    else if($h.EmailAddressDetails["isInstanceOfType"](recipients[i]))
                    {
                        var address = recipients[i];
                        $h.ComposeRecipient._throwOnInvalidDisplayNameOrEmail$p(address["displayName"],address["emailAddress"]);
                        recipientData[i] = $h.ComposeRecipient._createEmailDictionaryForHost$p(address["emailAddress"],address["displayName"])
                    }
                    else if(Object["isInstanceOfType"](recipients[i]))
                    {
                        var input = recipients[i];
                        var emailAddress = input["emailAddress"];
                        var displayName = input["displayName"];
                        $h.ComposeRecipient._throwOnInvalidDisplayNameOrEmail$p(displayName,emailAddress);
                        recipientData[i] = $h.ComposeRecipient._createEmailDictionaryForHost$p(emailAddress,displayName)
                    }
                    else
                        throw Error.argument("recipients");
            else
                throw Error.argument("recipients");
            var $$t_B = this;
            window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(isSet ? 21 : 22,{
                recipientField: this._type$p$0,
                recipientArray: recipientData
            },function(rawInput)
            {
                return rawInput
            },parameters._asyncContext$p$0,parameters._callback$p$0)
        }
    };
    $h.ComposeRecipient.prototype.getAsync = function()
    {
        var args = [];
        for(var $$pai_2 = 0; $$pai_2 < arguments["length"]; ++$$pai_2)
            args[$$pai_2] = arguments[$$pai_2];
        var parameters = $h.CommonParameters.parse(args,true);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,this._propertyName$p$0 + ".getAsync");
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(15,{recipientField: this._type$p$0},$h.ComposeRecipient._getAsyncFormatter$p,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ComposeRecipient.prototype.setAsync = function(recipients)
    {
        var args = [];
        for(var $$pai_2 = 1; $$pai_2 < arguments["length"]; ++$$pai_2)
            args[$$pai_2 - 1] = arguments[$$pai_2];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,this._propertyName$p$0 + ".setAsync");
        this.setAddHelper(recipients,args,true)
    };
    $h.ComposeRecipient.prototype.addAsync = function(recipients)
    {
        var args = [];
        for(var $$pai_2 = 1; $$pai_2 < arguments["length"]; ++$$pai_2)
            args[$$pai_2 - 1] = arguments[$$pai_2];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,this._propertyName$p$0 + ".addAsync");
        this.setAddHelper(recipients,args,false)
    };
    $h.ComposeRecipient.RecipientField = function(){};
    $h.ComposeRecipient.RecipientField.prototype = {
        to: 0,
        cc: 1,
        bcc: 2,
        requiredAttendees: 0,
        optionalAttendees: 1
    };
    $h.ComposeRecipient.RecipientField["registerEnum"]("$h.ComposeRecipient.RecipientField",false);
    $h.ComposeRecurrence = function(isInstance)
    {
        this._isInstance$p$0 = isInstance
    };
    $h.ComposeRecurrence.copyRecurrenceObjectConvertSeriesTimeJson = function(recurrenceObject)
    {
        var seriesTime = new window["Microsoft"]["Office"]["WebExtension"]["SeriesTime"];
        var recurrenceDictionary = recurrenceObject;
        var recurrenceCopy = {};
        if($h.ScriptHelpers.isNullOrUndefined(recurrenceDictionary["recurrenceProperties"]))
            recurrenceCopy["recurrenceProperties"] = null;
        else
            recurrenceCopy["recurrenceProperties"] = $h.ScriptHelpers.deepClone(recurrenceDictionary["recurrenceProperties"]);
        recurrenceCopy["recurrenceType"] = recurrenceDictionary["recurrenceType"];
        if($h.ScriptHelpers.isNullOrUndefined(recurrenceDictionary["recurrenceTimeZone"]))
            recurrenceCopy["recurrenceTimeZone"] = null;
        else
            recurrenceCopy["recurrenceTimeZone"] = $h.ScriptHelpers.deepClone(recurrenceDictionary["recurrenceTimeZone"]);
        seriesTime.importFromSeriesTimeJsonObject(recurrenceDictionary["seriesTimeJson"]);
        recurrenceCopy["seriesTime"] = seriesTime;
        return recurrenceCopy
    };
    $h.ComposeRecurrence._throwOnNullParameter$p = function(recurrenceObject, parameterName)
    {
        var recurrenceDictionary = recurrenceObject;
        if(!recurrenceDictionary[parameterName])
            throw Error.argumentNull(parameterName);
    };
    $h.ComposeRecurrence._throwOnInvalidRecurrenceType$p = function(recurrenceType)
    {
        if(recurrenceType !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecurrenceType"]["Daily"] && recurrenceType !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecurrenceType"]["Weekly"] && recurrenceType !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecurrenceType"]["Weekday"] && recurrenceType !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecurrenceType"]["Yearly"] && recurrenceType !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecurrenceType"]["Monthly"])
            throw Error.argument("recurrenceType");
    };
    $h.ComposeRecurrence._throwOnInvalidDailyRecurrence$p = function(recurrenceProperties)
    {
        $h.ComposeRecurrence._throwOnNullParameter$p(recurrenceProperties,"interval");
        window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(recurrenceProperties["interval"],Number,"interval")
    };
    $h.ComposeRecurrence._verifyDays$p = function(dayEnum, checkGroupedDays)
    {
        var fRegularDay = dayEnum === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Days"]["Mon"] || dayEnum === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Days"]["Tue"] || dayEnum === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Days"]["Wed"] || dayEnum === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Days"]["Thu"] || dayEnum === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Days"]["Fri"] || dayEnum === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Days"]["Sat"] || dayEnum === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Days"]["Sun"];
        if(checkGroupedDays)
        {
            var fGroupedDay = dayEnum === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Days"]["WeekendDay"] || dayEnum === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Days"]["Weekday"] || dayEnum === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Days"]["Day"];
            return fGroupedDay || fRegularDay
        }
        else
            return fRegularDay
    };
    $h.ComposeRecurrence._throwOnInvalidDaysArray$p = function(daysArray)
    {
        for(var i = 0; i < daysArray["length"]; i++)
            if(!$h.ComposeRecurrence._verifyDays$p(daysArray[i],false))
                throw Error.argument("days");
    };
    $h.ComposeRecurrence._throwOnInvalidWeeklyRecurrence$p = function(recurrenceProperties)
    {
        var recurrenceDictionary = recurrenceProperties;
        $h.ComposeRecurrence._throwOnNullParameter$p(recurrenceProperties,"interval");
        window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(recurrenceDictionary["interval"],Number,"interval");
        $h.ComposeRecurrence._throwOnNullParameter$p(recurrenceProperties,"days");
        window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(recurrenceDictionary["days"],Array,"days");
        $h.ComposeRecurrence._throwOnInvalidDaysArray$p(recurrenceDictionary["days"]);
        if(recurrenceDictionary["firstDayOfWeek"])
        {
            window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(recurrenceDictionary["firstDayOfWeek"],String,"firstDayOfWeek");
            if(!$h.ComposeRecurrence._verifyDays$p(recurrenceDictionary["firstDayOfWeek"],false))
                throw Error.argument("firstDayOfWeek");
        }
    };
    $h.ComposeRecurrence._throwOnInvalidWeekNumber$p = function(weekNumber)
    {
        if(weekNumber !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["WeekNumber"]["First"] && weekNumber !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["WeekNumber"]["Second"] && weekNumber !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["WeekNumber"]["Third"] && weekNumber !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["WeekNumber"]["Fourth"] && weekNumber !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["WeekNumber"]["Last"])
            throw Error.argument("weekNumber");
    };
    $h.ComposeRecurrence._throwOnInvalidDayOfMonth$p = function(iDayOfMonth)
    {
        if(iDayOfMonth < 1 || iDayOfMonth > 31)
            throw Error.argument("dayOfMonth");
    };
    $h.ComposeRecurrence._throwOnInvalidMonthlyRecurrence$p = function(recurrenceProperties)
    {
        var recurrenceDictionary = recurrenceProperties;
        $h.ComposeRecurrence._throwOnNullParameter$p(recurrenceProperties,"interval");
        window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(recurrenceDictionary["interval"],Number,"interval");
        if(recurrenceDictionary["dayOfMonth"])
        {
            window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(recurrenceDictionary["dayOfMonth"],Number,"dayOfMonth");
            $h.ComposeRecurrence._throwOnInvalidDayOfMonth$p(recurrenceDictionary["dayOfMonth"])
        }
        else if(recurrenceDictionary["dayOfWeek"] && recurrenceDictionary["weekNumber"])
        {
            window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(recurrenceDictionary["dayOfWeek"],String,"dayOfMonth");
            if(!$h.ComposeRecurrence._verifyDays$p(recurrenceDictionary["dayOfWeek"],true))
                throw Error.argument("dayOfWeek");
            window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(recurrenceDictionary["weekNumber"],String,"dayOfMonth");
            $h.ComposeRecurrence._throwOnInvalidWeekNumber$p(recurrenceDictionary["weekNumber"])
        }
        else
            throw Error.create(window["_u"]["ExtensibilityStrings"]["l_Recurrence_Error_Properties_Invalid_Text"]);
    };
    $h.ComposeRecurrence._throwOnInvalidMonth$p = function(month)
    {
        if(month !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Month"]["Jan"] && month !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Month"]["Feb"] && month !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Month"]["Mar"] && month !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Month"]["Apr"] && month !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Month"]["May"] && month !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Month"]["Jun"] && month !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Month"]["Jul"] && month !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Month"]["Aug"] && month !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Month"]["Sep"] && month !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Month"]["Oct"] && month !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Month"]["Nov"] && month !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Month"]["Dec"])
            throw Error.argument("month");
    };
    $h.ComposeRecurrence._throwOnInvalidYearlyRecurrence$p = function(recurrenceProperties)
    {
        var recurrenceDictionary = recurrenceProperties;
        $h.ComposeRecurrence._throwOnNullParameter$p(recurrenceProperties,"interval");
        window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(recurrenceDictionary["interval"],Number,"interval");
        $h.ComposeRecurrence._throwOnNullParameter$p(recurrenceProperties,"month");
        window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(recurrenceDictionary["month"],String,"month");
        $h.ComposeRecurrence._throwOnInvalidMonth$p(recurrenceDictionary["month"]);
        if(recurrenceDictionary["dayOfMonth"])
        {
            window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(recurrenceDictionary["dayOfMonth"],Number,"dayOfMonth");
            $h.ComposeRecurrence._throwOnInvalidDayOfMonth$p(recurrenceDictionary["dayOfMonth"])
        }
        else if(recurrenceDictionary["weekNumber"] && recurrenceDictionary["dayOfWeek"])
        {
            window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(recurrenceDictionary["dayOfWeek"],String,"dayOfMonth");
            if(!$h.ComposeRecurrence._verifyDays$p(recurrenceDictionary["dayOfWeek"],true))
                throw Error.argument("dayOfWeek");
            window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(recurrenceDictionary["weekNumber"],String,"dayOfMonth");
            $h.ComposeRecurrence._throwOnInvalidWeekNumber$p(recurrenceDictionary["weekNumber"])
        }
        else
            throw Error.create(window["_u"]["ExtensibilityStrings"]["l_Recurrence_Error_Properties_Invalid_Text"]);
    };
    $h.ComposeRecurrence.verifyRecurrenceObject = function(recurrenceObject)
    {
        if(!recurrenceObject)
            return;
        var recurrenceDictionary = recurrenceObject;
        $h.ComposeRecurrence._throwOnNullParameter$p(recurrenceObject,"recurrenceType");
        $h.ComposeRecurrence._throwOnNullParameter$p(recurrenceObject,"seriesTime");
        if(!window["Microsoft"]["Office"]["WebExtension"]["SeriesTime"]["isInstanceOfType"](recurrenceDictionary["seriesTime"]) || !recurrenceDictionary["seriesTime"].isValid())
            throw Error.argument("seriesTime");
        if(!recurrenceDictionary["seriesTime"].isEndAfterStart())
            throw Error.create(window["_u"]["ExtensibilityStrings"]["l_InvalidEventDates_Text"]);
        $h.ComposeRecurrence._throwOnInvalidRecurrenceType$p(recurrenceDictionary["recurrenceType"]);
        if(recurrenceDictionary["recurrenceType"] !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecurrenceType"]["Weekday"])
            $h.ComposeRecurrence._throwOnNullParameter$p(recurrenceObject,"recurrenceProperties");
        if(recurrenceDictionary["recurrenceTimeZone"])
        {
            window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(recurrenceDictionary["recurrenceTimeZone"],Object,"recurrenceTimeZone");
            var recurrenceTimeZone = recurrenceDictionary["recurrenceTimeZone"];
            $h.ComposeRecurrence._throwOnNullParameter$p(recurrenceTimeZone,"name");
            window["OSF"]["DDA"]["OutlookAppOm"].throwOnArgumentType(recurrenceTimeZone["name"],String,"name")
        }
        if(recurrenceDictionary["recurrenceType"] === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecurrenceType"]["Daily"])
            $h.ComposeRecurrence._throwOnInvalidDailyRecurrence$p(recurrenceDictionary["recurrenceProperties"]);
        else if(recurrenceDictionary["recurrenceType"] === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecurrenceType"]["Weekly"])
            $h.ComposeRecurrence._throwOnInvalidWeeklyRecurrence$p(recurrenceDictionary["recurrenceProperties"]);
        else if(recurrenceDictionary["recurrenceType"] === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecurrenceType"]["Monthly"])
            $h.ComposeRecurrence._throwOnInvalidMonthlyRecurrence$p(recurrenceDictionary["recurrenceProperties"]);
        else if(recurrenceDictionary["recurrenceType"] === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecurrenceType"]["Yearly"])
            $h.ComposeRecurrence._throwOnInvalidYearlyRecurrence$p(recurrenceDictionary["recurrenceProperties"])
    };
    $h.ComposeRecurrence.prototype = {
        _isInstance$p$0: false,
        convertSeriesTime: function(recurrenceObject)
        {
            var recurrenceDictionary = recurrenceObject;
            if(recurrenceDictionary && recurrenceDictionary["seriesTime"])
                if(window["Microsoft"]["Office"]["WebExtension"]["SeriesTime"]["isInstanceOfType"](recurrenceDictionary["seriesTime"]))
                {
                    var recurrenceCopy = {};
                    if($h.ScriptHelpers.isNullOrUndefined(recurrenceDictionary["recurrenceProperties"]))
                        recurrenceCopy["recurrenceProperties"] = null;
                    else
                        recurrenceCopy["recurrenceProperties"] = $h.ScriptHelpers.deepClone(recurrenceDictionary["recurrenceProperties"]);
                    recurrenceCopy["recurrenceType"] = recurrenceDictionary["recurrenceType"];
                    if($h.ScriptHelpers.isNullOrUndefined(recurrenceDictionary["recurrenceTimeZone"]))
                        recurrenceCopy["recurrenceTimeZone"] = null;
                    else
                        recurrenceCopy["recurrenceTimeZone"] = $h.ScriptHelpers.deepClone(recurrenceDictionary["recurrenceTimeZone"]);
                    recurrenceCopy["seriesTimeJson"] = recurrenceDictionary["seriesTime"].exportToSeriesTimeJsonDictionary();
                    return recurrenceCopy
                }
            return recurrenceObject
        }
    };
    $h.ComposeRecurrence.prototype.getAsync = function()
    {
        var args = [];
        for(var $$pai_2 = 0; $$pai_2 < arguments["length"]; ++$$pai_2)
            args[$$pai_2] = arguments[$$pai_2];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"recurrence.getAsync");
        var parameters = $h.CommonParameters.parse(args,true);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(103,null,window["Microsoft"]["Office"]["WebExtension"]["OutlookBase"]["SeriesTimeJsonConverter"],parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ComposeRecurrence.prototype.setAsync = function(recurrenceObject)
    {
        var args = [];
        for(var $$pai_3 = 1; $$pai_3 < arguments["length"]; ++$$pai_3)
            args[$$pai_3 - 1] = arguments[$$pai_3];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"recurrence.setAsync");
        if(this._isInstance$p$0)
            throw Error.create(window["_u"]["ExtensibilityStrings"]["l_Recurrence_Error_Instance_SetAsync_Text"]);
        $h.ComposeRecurrence.verifyRecurrenceObject(recurrenceObject);
        var parameters = $h.CommonParameters.parse(args,false);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(104,{recurrenceData: this.convertSeriesTime(recurrenceObject)},null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ComposeLocation = function(){};
    $h.ComposeLocation.prototype.getAsync = function()
    {
        var args = [];
        for(var $$pai_2 = 0; $$pai_2 < arguments["length"]; ++$$pai_2)
            args[$$pai_2] = arguments[$$pai_2];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"location.getAsync");
        var parameters = $h.CommonParameters.parse(args,true);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(26,null,null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ComposeLocation.prototype.setAsync = function(location)
    {
        var args = [];
        for(var $$pai_3 = 1; $$pai_3 < arguments["length"]; ++$$pai_3)
            args[$$pai_3 - 1] = arguments[$$pai_3];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"location.setAsync");
        var parameters = $h.CommonParameters.parse(args,false);
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(location.length,0,255,"location");
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(27,{location: location},null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ComposeSubject = function(){};
    $h.ComposeSubject.prototype.getAsync = function()
    {
        var args = [];
        for(var $$pai_2 = 0; $$pai_2 < arguments["length"]; ++$$pai_2)
            args[$$pai_2] = arguments[$$pai_2];
        var parameters = $h.CommonParameters.parse(args,true);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"subject.getAsync");
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(18,null,null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ComposeSubject.prototype.setAsync = function(data)
    {
        var args = [];
        for(var $$pai_3 = 1; $$pai_3 < arguments["length"]; ++$$pai_3)
            args[$$pai_3 - 1] = arguments[$$pai_3];
        var parameters = $h.CommonParameters.parse(args,false);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"subject.setAsync");
        if(!String["isInstanceOfType"](data))
            throw Error.argument("data");
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(data.length,0,255,"data");
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(17,{subject: data},null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ComposeTime = function(type)
    {
        this.$$d__ticksToDateFormatter$p$0 = Function.createDelegate(this,this._ticksToDateFormatter$p$0);
        this._timeType$p$0 = type
    };
    $h.ComposeTime.prototype = {
        _timeType$p$0: 0,
        _ticksToDateFormatter$p$0: function(rawInput)
        {
            var ticks = rawInput;
            return new Date(ticks)
        },
        _getPropertyName$p$0: function()
        {
            return this._timeType$p$0 === 1 ? "start" : "end"
        }
    };
    $h.ComposeTime.prototype.getAsync = function()
    {
        var args = [];
        for(var $$pai_2 = 0; $$pai_2 < arguments["length"]; ++$$pai_2)
            args[$$pai_2] = arguments[$$pai_2];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,this._getPropertyName$p$0() + ".getAsync");
        var parameters = $h.CommonParameters.parse(args,true);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(24,{TimeProperty: this._timeType$p$0},this.$$d__ticksToDateFormatter$p$0,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ComposeTime.prototype.setAsync = function(dateTime)
    {
        var args = [];
        for(var $$pai_3 = 1; $$pai_3 < arguments["length"]; ++$$pai_3)
            args[$$pai_3 - 1] = arguments[$$pai_3];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,this._getPropertyName$p$0() + ".setAsync");
        if(!Date["isInstanceOfType"](dateTime))
            throw Error.argumentType("dateTime",Object["getType"](dateTime),Date);
        if(window["isNaN"](dateTime["getTime"]()))
            throw Error.argument("dateTime");
        if(dateTime["getTime"]() < -864e13 || dateTime["getTime"]() > 864e13)
            throw Error.argumentOutOfRange("dateTime");
        var parameters = $h.CommonParameters.parse(args,false);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(25,{
            TimeProperty: this._timeType$p$0,
            time: dateTime["getTime"]()
        },null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ComposeTime.TimeType = function(){};
    $h.ComposeTime.TimeType.prototype = {
        start: 1,
        end: 2
    };
    $h.ComposeTime.TimeType["registerEnum"]("$h.ComposeTime.TimeType",false);
    $h.Contact = function(data)
    {
        this.$$d__getContactString$p$0 = Function.createDelegate(this,this._getContactString$p$0);
        this.$$d__getAddresses$p$0 = Function.createDelegate(this,this._getAddresses$p$0);
        this.$$d__getUrls$p$0 = Function.createDelegate(this,this._getUrls$p$0);
        this.$$d__getEmailAddresses$p$0 = Function.createDelegate(this,this._getEmailAddresses$p$0);
        this.$$d__getPhoneNumbers$p$0 = Function.createDelegate(this,this._getPhoneNumbers$p$0);
        this.$$d__getBusinessName$p$0 = Function.createDelegate(this,this._getBusinessName$p$0);
        this.$$d__getPersonName$p$0 = Function.createDelegate(this,this._getPersonName$p$0);
        this._data$p$0 = data;
        $h.InitialData._defineReadOnlyProperty$i(this,"personName",this.$$d__getPersonName$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"businessName",this.$$d__getBusinessName$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"phoneNumbers",this.$$d__getPhoneNumbers$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"emailAddresses",this.$$d__getEmailAddresses$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"urls",this.$$d__getUrls$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"addresses",this.$$d__getAddresses$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"contactString",this.$$d__getContactString$p$0)
    };
    $h.Contact.prototype = {
        _data$p$0: null,
        _phoneNumbers$p$0: null,
        _getPersonName$p$0: function()
        {
            return this._data$p$0["PersonName"]
        },
        _getBusinessName$p$0: function()
        {
            return this._data$p$0["BusinessName"]
        },
        _getAddresses$p$0: function()
        {
            return $h.Entities._getExtractedStringProperty$i(this._data$p$0,"Addresses")
        },
        _getEmailAddresses$p$0: function()
        {
            return $h.Entities._getExtractedStringProperty$i(this._data$p$0,"EmailAddresses")
        },
        _getUrls$p$0: function()
        {
            return $h.Entities._getExtractedStringProperty$i(this._data$p$0,"Urls")
        },
        _getPhoneNumbers$p$0: function()
        {
            if(!this._phoneNumbers$p$0)
            {
                var $$t_1 = this;
                this._phoneNumbers$p$0 = $h.Entities._getExtractedObjects$i($h.PhoneNumber,this._data$p$0,"PhoneNumbers",function(data)
                {
                    return new $h.PhoneNumber(data)
                })
            }
            return this._phoneNumbers$p$0
        },
        _getContactString$p$0: function()
        {
            return this._data$p$0["ContactString"]
        }
    };
    $h.CustomProperties = function(data)
    {
        if($h.ScriptHelpers.isNullOrUndefined(data))
            throw Error.argumentNull("data");
        if(Array["isInstanceOfType"](data))
        {
            var customPropertiesArray = data;
            if(customPropertiesArray["length"] > 0)
                this._data$p$0 = customPropertiesArray[0];
            else
                throw Error.argument("data");
        }
        else
            this._data$p$0 = data
    };
    $h.CustomProperties.prototype = {_data$p$0: null};
    $h.CustomProperties.prototype.get = function(name)
    {
        var value = this._data$p$0[name];
        if(typeof value === "string")
        {
            var valueString = value;
            if(valueString.length > 6 && valueString.startsWith("Date(") && valueString.endsWith(")"))
            {
                var ticksString = valueString.substring(5,valueString.length - 1);
                var ticks = window["parseInt"](ticksString);
                if(!window["isNaN"](ticks))
                {
                    var dateTimeValue = new Date(ticks);
                    if(dateTimeValue)
                        value = dateTimeValue
                }
            }
        }
        return value
    };
    $h.CustomProperties.prototype.set = function(name, value)
    {
        if(window["OSF"]["OUtil"]["isDate"](value))
            value = "Date(" + value["getTime"]() + ")";
        this._data$p$0[name] = value
    };
    $h.CustomProperties.prototype.remove = function(name)
    {
        delete this._data$p$0[name]
    };
    $h.CustomProperties.prototype.saveAsync = function()
    {
        var args = [];
        for(var $$pai_4 = 0; $$pai_4 < arguments["length"]; ++$$pai_4)
            args[$$pai_4] = arguments[$$pai_4];
        var MaxCustomPropertiesLength = 2500;
        if(window["JSON"]["stringify"](this._data$p$0).length > MaxCustomPropertiesLength)
            throw Error.argument();
        var parameters = $h.CommonParameters.parse(args,false,true);
        if(window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.shouldRunNewCode($h.ShouldRunNewCodeForFlags.saveCustomProperties))
            window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(4,{customProperties: this._data$p$0},null,parameters._asyncContext$p$0,parameters._callback$p$0);
        else
        {
            var saveCustomProperties = new $h.SaveDictionaryRequest(parameters._callback$p$0,parameters._asyncContext$p$0);
            saveCustomProperties._sendRequest$i$0(4,"SaveCustomProperties",{customProperties: this._data$p$0})
        }
    };
    $h.Diagnostics = function(data, appName)
    {
        this.$$d__getOwaView$p$0 = Function.createDelegate(this,this._getOwaView$p$0);
        this.$$d__getHostVersion$p$0 = Function.createDelegate(this,this._getHostVersion$p$0);
        this.$$d__getHostName$p$0 = Function.createDelegate(this,this._getHostName$p$0);
        this._data$p$0 = data;
        this._appName$p$0 = appName;
        $h.InitialData._defineReadOnlyProperty$i(this,"hostName",this.$$d__getHostName$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"hostVersion",this.$$d__getHostVersion$p$0);
        if(64 === this._appName$p$0)
            $h.InitialData._defineReadOnlyProperty$i(this,"OWAView",this.$$d__getOwaView$p$0)
    };
    $h.Diagnostics.prototype = {
        _data$p$0: null,
        _appName$p$0: 0,
        _getHostName$p$0: function()
        {
            switch(this._appName$p$0)
            {
                case 8:
                    return"Outlook";
                case 64:
                    return"OutlookWebApp";
                case 65536:
                    return"OutlookIOS";
                case 4194304:
                    return"OutlookAndroid";
                default:
                    return null
            }
        },
        _getHostVersion$p$0: function()
        {
            return this._data$p$0.get__hostVersion$i$0()
        },
        _getOwaView$p$0: function()
        {
            return this._data$p$0.get__owaView$i$0()
        }
    };
    $h.EmailAddressDetails = function(data)
    {
        this.$$d__getRecipientType$p$0 = Function.createDelegate(this,this._getRecipientType$p$0);
        this.$$d__getAppointmentResponse$p$0 = Function.createDelegate(this,this._getAppointmentResponse$p$0);
        this.$$d__getDisplayName$p$0 = Function.createDelegate(this,this._getDisplayName$p$0);
        this.$$d__getEmailAddress$p$0 = Function.createDelegate(this,this._getEmailAddress$p$0);
        this._data$p$0 = data;
        $h.InitialData._defineReadOnlyProperty$i(this,"emailAddress",this.$$d__getEmailAddress$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"displayName",this.$$d__getDisplayName$p$0);
        if($h.ScriptHelpers.dictionaryContainsKey(data,"appointmentResponse"))
            $h.InitialData._defineReadOnlyProperty$i(this,"appointmentResponse",this.$$d__getAppointmentResponse$p$0);
        if($h.ScriptHelpers.dictionaryContainsKey(data,"recipientType"))
            $h.InitialData._defineReadOnlyProperty$i(this,"recipientType",this.$$d__getRecipientType$p$0)
    };
    $h.EmailAddressDetails._createFromEmailUserDictionary$i = function(data)
    {
        var emailAddressDetailsDictionary = {};
        var displayName = data["Name"];
        var emailAddress = data["UserId"];
        emailAddressDetailsDictionary["name"] = displayName || $h.EmailAddressDetails._emptyString$p;
        emailAddressDetailsDictionary["address"] = emailAddress || $h.EmailAddressDetails._emptyString$p;
        return new $h.EmailAddressDetails(emailAddressDetailsDictionary)
    };
    $h.EmailAddressDetails.prototype = {
        _data$p$0: null,
        _getEmailAddress$p$0: function()
        {
            return this._data$p$0["address"]
        },
        _getDisplayName$p$0: function()
        {
            return this._data$p$0["name"]
        },
        _getAppointmentResponse$p$0: function()
        {
            var response = this._data$p$0["appointmentResponse"];
            return response < $h.EmailAddressDetails._responseTypeMap$p["length"] ? $h.EmailAddressDetails._responseTypeMap$p[response] : window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ResponseType"]["None"]
        },
        _getRecipientType$p$0: function()
        {
            var response = this._data$p$0["recipientType"];
            return response < $h.EmailAddressDetails._recipientTypeMap$p["length"] ? $h.EmailAddressDetails._recipientTypeMap$p[response] : window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecipientType"]["Other"]
        }
    };
    $h.EmailAddressDetails.prototype.toJSON = function()
    {
        var result = {};
        result["emailAddress"] = this._getEmailAddress$p$0();
        result["displayName"] = this._getDisplayName$p$0();
        if($h.ScriptHelpers.dictionaryContainsKey(this._data$p$0,"appointmentResponse"))
            result["appointmentResponse"] = this._getAppointmentResponse$p$0();
        if($h.ScriptHelpers.dictionaryContainsKey(this._data$p$0,"recipientType"))
            result["recipientType"] = this._getRecipientType$p$0();
        return result
    };
    $h.EnhancedLocation = function(supportsWriteMethods)
    {
        this.$$d_getAsyncApi = Function.createDelegate(this,this.getAsyncApi);
        this.$$d_removeAsyncApi = Function.createDelegate(this,this.removeAsyncApi);
        this.$$d_addAsyncApi = Function.createDelegate(this,this.addAsyncApi);
        var currentInstance = this;
        if(supportsWriteMethods)
        {
            currentInstance["addAsync"] = this.$$d_addAsyncApi;
            currentInstance["removeAsync"] = this.$$d_removeAsyncApi
        }
        currentInstance["getAsync"] = this.$$d_getAsyncApi
    };
    $h.EnhancedLocation.prototype = {
        getAsyncApi: function()
        {
            var args = [];
            for(var $$pai_2 = 0; $$pai_2 < arguments["length"]; ++$$pai_2)
                args[$$pai_2] = arguments[$$pai_2];
            window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"enhancedLocation.getAsync");
            var parameters = $h.CommonParameters.parse(args,true);
            window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(154,null,null,parameters._asyncContext$p$0,parameters._callback$p$0)
        },
        addAsyncApi: function(locationIdentifiersArray)
        {
            var args = [];
            for(var $$pai_3 = 1; $$pai_3 < arguments["length"]; ++$$pai_3)
                args[$$pai_3 - 1] = arguments[$$pai_3];
            window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"enhancedLocation.addAsync");
            var parameters = $h.CommonParameters.parse(args,false);
            this._validateLocationIdentifiers$p$0(locationIdentifiersArray);
            window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(155,{enhancedLocations: locationIdentifiersArray},null,parameters._asyncContext$p$0,parameters._callback$p$0)
        },
        removeAsyncApi: function(locationIdentifiersArray)
        {
            var args = [];
            for(var $$pai_3 = 1; $$pai_3 < arguments["length"]; ++$$pai_3)
                args[$$pai_3 - 1] = arguments[$$pai_3];
            window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(2,"enhancedLocation.removeAsync");
            var parameters = $h.CommonParameters.parse(args,false);
            this._validateLocationIdentifiers$p$0(locationIdentifiersArray);
            window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(156,{enhancedLocations: locationIdentifiersArray},null,parameters._asyncContext$p$0,parameters._callback$p$0)
        },
        _validateLocationIdentifiers$p$0: function(locationIdentifier)
        {
            if($h.ScriptHelpers.isNullOrUndefined(locationIdentifier))
                throw Error.argument("locationIdentifier");
            if(!Array["isInstanceOfType"](locationIdentifier))
                throw Error.argumentType("locationIdentifier",Object["getType"](locationIdentifier),Array);
            if(!locationIdentifier["length"])
                throw Error.argument("locationIdentifier");
            for(var $$arr_1 = locationIdentifier, $$len_2 = $$arr_1.length, $$idx_3 = 0; $$idx_3 < $$len_2; ++$$idx_3)
            {
                var locationIdentifiersDictionary = $$arr_1[$$idx_3];
                this._validateLocationIdentifierDictionary$p$0(locationIdentifiersDictionary);
                var locationIdentifierType = locationIdentifiersDictionary["type"];
                if(locationIdentifierType === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["LocationType"]["Room"] || locationIdentifierType === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["LocationType"]["Custom"])
                    this._validateIdParameter$p$0(locationIdentifiersDictionary["id"],locationIdentifierType);
                else
                    throw Error.argument("type");
            }
        },
        _validateIdParameter$p$0: function(id, type)
        {
            if(!$h.ScriptHelpers.isNonEmptyString(id))
                throw Error.argument("id");
            if(type === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["LocationType"]["Room"])
                if(id.length > 571)
                    throw Error.argument("id");
        },
        _validateLocationIdentifierDictionary$p$0: function(dict)
        {
            if($h.ScriptHelpers.isNullOrUndefined(dict))
                throw Error.argument("locationIdentifier");
            var keys = Object["keys"](dict);
            if(keys["length"] !== 2)
                throw Error.argument("locationIdentifier");
        }
    };
    $h.Entities = function(data, filteredEntitiesData, timeSent, permissionLevel)
    {
        this.$$d__createMeetingSuggestion$p$0 = Function.createDelegate(this,this._createMeetingSuggestion$p$0);
        this.$$d__getParcelDeliveries$p$0 = Function.createDelegate(this,this._getParcelDeliveries$p$0);
        this.$$d__getFlightReservations$p$0 = Function.createDelegate(this,this._getFlightReservations$p$0);
        this.$$d__getContacts$p$0 = Function.createDelegate(this,this._getContacts$p$0);
        this.$$d__getPhoneNumbers$p$0 = Function.createDelegate(this,this._getPhoneNumbers$p$0);
        this.$$d__getUrls$p$0 = Function.createDelegate(this,this._getUrls$p$0);
        this.$$d__getEmailAddresses$p$0 = Function.createDelegate(this,this._getEmailAddresses$p$0);
        this.$$d__getMeetingSuggestions$p$0 = Function.createDelegate(this,this._getMeetingSuggestions$p$0);
        this.$$d__getTaskSuggestions$p$0 = Function.createDelegate(this,this._getTaskSuggestions$p$0);
        this.$$d__getAddresses$p$0 = Function.createDelegate(this,this._getAddresses$p$0);
        this._data$p$0 = data || {};
        this._filteredData$p$0 = filteredEntitiesData || {};
        this._dateTimeSent$p$0 = timeSent;
        $h.InitialData._defineReadOnlyProperty$i(this,"addresses",this.$$d__getAddresses$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"taskSuggestions",this.$$d__getTaskSuggestions$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"meetingSuggestions",this.$$d__getMeetingSuggestions$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"emailAddresses",this.$$d__getEmailAddresses$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"urls",this.$$d__getUrls$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"phoneNumbers",this.$$d__getPhoneNumbers$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"contacts",this.$$d__getContacts$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"flightReservations",this.$$d__getFlightReservations$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"parcelDeliveries",this.$$d__getParcelDeliveries$p$0);
        this._permissionLevel$p$0 = permissionLevel
    };
    $h.Entities._getExtractedObjects$i = function(T, data, name, creator, removeDuplicates, stringPropertyName)
    {
        var results = null;
        var extractedObjects = data[name];
        if(!extractedObjects)
            return new Array(0);
        if(removeDuplicates)
            extractedObjects = $h.Entities._removeDuplicate$p(Object,extractedObjects,$h.Entities._entityDictionaryEquals$p,stringPropertyName);
        results = new Array(extractedObjects["length"]);
        var count = 0;
        for(var $$arr_9 = extractedObjects, $$len_A = $$arr_9.length, $$idx_B = 0; $$idx_B < $$len_A; ++$$idx_B)
        {
            var extractedObject = $$arr_9[$$idx_B];
            if(name === "MeetingSuggestions")
                extractedObject["IsLegacyEntityExtraction"] = "IsLegacyEntityExtraction" in data ? data["IsLegacyEntityExtraction"] : true;
            if(creator)
                results[count++] = creator(extractedObject);
            else
                results[count++] = extractedObject
        }
        return results
    };
    $h.Entities._getExtractedStringProperty$i = function(data, name, removeDuplicate)
    {
        var extractedProperties = data[name];
        if(!extractedProperties)
            return new Array(0);
        if(removeDuplicate)
            extractedProperties = $h.Entities._removeDuplicate$p(String,extractedProperties,$h.Entities._stringEquals$p,null);
        return extractedProperties
    };
    $h.Entities._createContact$p = function(data)
    {
        return new $h.Contact(data)
    };
    $h.Entities._createTaskSuggestion$p = function(data)
    {
        return new $h.TaskSuggestion(data)
    };
    $h.Entities._createPhoneNumber$p = function(data)
    {
        return new $h.PhoneNumber(data)
    };
    $h.Entities._entityDictionaryEquals$p = function(dictionary1, dictionary2, entityPropertyIdentifier)
    {
        if(dictionary1 === dictionary2)
            return true;
        if(!dictionary1 || !dictionary2)
            return false;
        if(dictionary1[entityPropertyIdentifier] === dictionary2[entityPropertyIdentifier])
            return true;
        return false
    };
    $h.Entities._stringEquals$p = function(string1, string2, entityProperty)
    {
        return string1 === string2
    };
    $h.Entities._removeDuplicate$p = function(T, array, entityEquals, entityPropertyIdentifier)
    {
        for(var matchIndex1 = array["length"] - 1; matchIndex1 >= 0; matchIndex1--)
        {
            var removeMatch = false;
            for(var matchIndex2 = matchIndex1 - 1; matchIndex2 >= 0; matchIndex2--)
                if(entityEquals(array[matchIndex1],array[matchIndex2],entityPropertyIdentifier))
                {
                    removeMatch = true;
                    break
                }
            if(removeMatch)
                Array.removeAt(array,matchIndex1)
        }
        return array
    };
    $h.Entities.prototype = {
        _dateTimeSent$p$0: null,
        _data$p$0: null,
        _filteredData$p$0: null,
        _filteredEntitiesCache$p$0: null,
        _permissionLevel$p$0: 0,
        _taskSuggestions$p$0: null,
        _meetingSuggestions$p$0: null,
        _phoneNumbers$p$0: null,
        _contacts$p$0: null,
        _addresses$p$0: null,
        _emailAddresses$p$0: null,
        _urls$p$0: null,
        _flightReservations$p$0: null,
        _parcelDeliveries$p$0: null,
        _getByType$i$0: function(entityType)
        {
            if(entityType === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["EntityType"]["MeetingSuggestion"])
                return this._getMeetingSuggestions$p$0();
            else if(entityType === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["EntityType"]["TaskSuggestion"])
                return this._getTaskSuggestions$p$0();
            else if(entityType === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["EntityType"]["Address"])
                return this._getAddresses$p$0();
            else if(entityType === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["EntityType"]["PhoneNumber"])
                return this._getPhoneNumbers$p$0();
            else if(entityType === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["EntityType"]["EmailAddress"])
                return this._getEmailAddresses$p$0();
            else if(entityType === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["EntityType"]["Url"])
                return this._getUrls$p$0();
            else if(entityType === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["EntityType"]["Contact"])
                return this._getContacts$p$0();
            else if(entityType === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["EntityType"]["FlightReservations"])
                return this._getFlightReservations$p$0();
            else if(entityType === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["EntityType"]["ParcelDeliveries"])
                return this._getParcelDeliveries$p$0();
            return null
        },
        _getFilteredEntitiesByName$i$0: function(name)
        {
            if(!this._filteredEntitiesCache$p$0)
                this._filteredEntitiesCache$p$0 = {};
            if(!$h.ScriptHelpers.dictionaryContainsKey(this._filteredEntitiesCache$p$0,name))
            {
                var found = false;
                for(var i = 0; i < $h.Entities._allEntityKeys$p["length"]; i++)
                {
                    var entityTypeKey = $h.Entities._allEntityKeys$p[i];
                    var perEntityTypeDictionary = this._filteredData$p$0[entityTypeKey];
                    if(!perEntityTypeDictionary)
                        continue;
                    if($h.ScriptHelpers.dictionaryContainsKey(perEntityTypeDictionary,name))
                    {
                        switch(entityTypeKey)
                        {
                            case"EmailAddresses":
                            case"Urls":
                                this._filteredEntitiesCache$p$0[name] = $h.Entities._getExtractedStringProperty$i(perEntityTypeDictionary,name);
                                break;
                            case"Addresses":
                                this._filteredEntitiesCache$p$0[name] = $h.Entities._getExtractedStringProperty$i(perEntityTypeDictionary,name,true);
                                break;
                            case"PhoneNumbers":
                                this._filteredEntitiesCache$p$0[name] = $h.Entities._getExtractedObjects$i($h.PhoneNumber,perEntityTypeDictionary,name,$h.Entities._createPhoneNumber$p,false,null);
                                break;
                            case"TaskSuggestions":
                                this._filteredEntitiesCache$p$0[name] = $h.Entities._getExtractedObjects$i($h.TaskSuggestion,perEntityTypeDictionary,name,$h.Entities._createTaskSuggestion$p,true,"TaskString");
                                break;
                            case"MeetingSuggestions":
                                this._filteredEntitiesCache$p$0[name] = $h.Entities._getExtractedObjects$i($h.MeetingSuggestion,perEntityTypeDictionary,name,this.$$d__createMeetingSuggestion$p$0,true,"MeetingString");
                                break;
                            case"Contacts":
                                this._filteredEntitiesCache$p$0[name] = $h.Entities._getExtractedObjects$i($h.Contact,perEntityTypeDictionary,name,$h.Entities._createContact$p,true,"ContactString");
                                break
                        }
                        found = true;
                        break
                    }
                }
                if(!found)
                    this._filteredEntitiesCache$p$0[name] = null
            }
            return this._filteredEntitiesCache$p$0[name]
        },
        _createMeetingSuggestion$p$0: function(data)
        {
            return new $h.MeetingSuggestion(data,this._dateTimeSent$p$0)
        },
        _getAddresses$p$0: function()
        {
            if(!this._addresses$p$0)
                this._addresses$p$0 = $h.Entities._getExtractedStringProperty$i(this._data$p$0,"Addresses",true);
            return this._addresses$p$0
        },
        _getEmailAddresses$p$0: function()
        {
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._permissionLevel$p$0);
            if(!this._emailAddresses$p$0)
                this._emailAddresses$p$0 = $h.Entities._getExtractedStringProperty$i(this._data$p$0,"EmailAddresses",false);
            return this._emailAddresses$p$0
        },
        _getUrls$p$0: function()
        {
            if(!this._urls$p$0)
                this._urls$p$0 = $h.Entities._getExtractedStringProperty$i(this._data$p$0,"Urls",false);
            return this._urls$p$0
        },
        _getPhoneNumbers$p$0: function()
        {
            if(!this._phoneNumbers$p$0)
                this._phoneNumbers$p$0 = $h.Entities._getExtractedObjects$i($h.PhoneNumber,this._data$p$0,"PhoneNumbers",$h.Entities._createPhoneNumber$p);
            return this._phoneNumbers$p$0
        },
        _getTaskSuggestions$p$0: function()
        {
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._permissionLevel$p$0);
            if(!this._taskSuggestions$p$0)
                this._taskSuggestions$p$0 = $h.Entities._getExtractedObjects$i($h.TaskSuggestion,this._data$p$0,"TaskSuggestions",$h.Entities._createTaskSuggestion$p,true,"TaskString");
            return this._taskSuggestions$p$0
        },
        _getMeetingSuggestions$p$0: function()
        {
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._permissionLevel$p$0);
            if(!this._meetingSuggestions$p$0)
                this._meetingSuggestions$p$0 = $h.Entities._getExtractedObjects$i($h.MeetingSuggestion,this._data$p$0,"MeetingSuggestions",this.$$d__createMeetingSuggestion$p$0,true,"MeetingString");
            return this._meetingSuggestions$p$0
        },
        _getContacts$p$0: function()
        {
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._permissionLevel$p$0);
            if(!this._contacts$p$0)
                this._contacts$p$0 = $h.Entities._getExtractedObjects$i($h.Contact,this._data$p$0,"Contacts",$h.Entities._createContact$p,true,"ContactString");
            return this._contacts$p$0
        },
        _getParcelDeliveries$p$0: function()
        {
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._permissionLevel$p$0);
            if(!this._parcelDeliveries$p$0)
                this._parcelDeliveries$p$0 = $h.Entities._getExtractedObjects$i(Object,this._data$p$0,"ParcelDeliveries",null);
            return this._parcelDeliveries$p$0
        },
        _getFlightReservations$p$0: function()
        {
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._permissionLevel$p$0);
            if(!this._flightReservations$p$0)
                this._flightReservations$p$0 = $h.Entities._getExtractedObjects$i(Object,this._data$p$0,"FlightReservations",null);
            return this._flightReservations$p$0
        }
    };
    window["Microsoft"]["Office"]["WebExtension"]["SeriesTime"] = function Microsoft_Office_WebExtension_SeriesTime()
    {
        this._startYear$p$0 = 0;
        this._startMonth$p$0 = 0;
        this._startDay$p$0 = 0;
        this._endYear$p$0 = 0;
        this._endMonth$p$0 = 0;
        this._endDay$p$0 = 0;
        this._startTimeMinutes$p$0 = 0;
        this._durationMinutes$p$0 = 0
    };
    Microsoft.Office.WebExtension.SeriesTime.prototype = {
        _startYear$p$0: 0,
        _startMonth$p$0: 0,
        _startDay$p$0: 0,
        _endYear$p$0: 0,
        _endMonth$p$0: 0,
        _endDay$p$0: 0,
        _startTimeMinutes$p$0: 0,
        _durationMinutes$p$0: 0,
        exportToSeriesTimeJsonDictionary: function()
        {
            var result = {};
            result["startYear"] = this._startYear$p$0;
            result["startMonth"] = this._startMonth$p$0;
            result["startDay"] = this._startDay$p$0;
            if(!this._endYear$p$0 && !this._endMonth$p$0 && !this._endDay$p$0)
                result["noEndDate"] = true;
            else
            {
                result["endYear"] = this._endYear$p$0;
                result["endMonth"] = this._endMonth$p$0;
                result["endDay"] = this._endDay$p$0
            }
            result["startTimeMin"] = this._startTimeMinutes$p$0;
            if(this._durationMinutes$p$0 > 0)
                result["durationMin"] = this._durationMinutes$p$0;
            return result
        },
        importFromSeriesTimeJsonObject: function(jsonObject)
        {
            var jsonDictionary = jsonObject;
            this._startYear$p$0 = jsonDictionary["startYear"];
            this._startMonth$p$0 = jsonDictionary["startMonth"];
            this._startDay$p$0 = jsonDictionary["startDay"];
            if(jsonDictionary["noEndDate"] && jsonDictionary["noEndDate"])
            {
                this._endYear$p$0 = 0;
                this._endMonth$p$0 = 0;
                this._endDay$p$0 = 0
            }
            else
            {
                this._endYear$p$0 = jsonDictionary["endYear"];
                this._endMonth$p$0 = jsonDictionary["endMonth"];
                this._endDay$p$0 = jsonDictionary["endDay"]
            }
            this._startTimeMinutes$p$0 = jsonDictionary["startTimeMin"];
            this._durationMinutes$p$0 = jsonDictionary["durationMin"]
        },
        isValid: function()
        {
            if(!this._isValidDate$p$0(this._startYear$p$0,this._startMonth$p$0,this._startDay$p$0))
                return false;
            if(this._endDay$p$0 && this._endMonth$p$0 && this._endYear$p$0)
                if(!this._isValidDate$p$0(this._endYear$p$0,this._endMonth$p$0,this._endDay$p$0))
                    return false;
            if(this._startTimeMinutes$p$0 < 0 || this._durationMinutes$p$0 <= 0)
                return false;
            return true
        },
        isEndAfterStart: function()
        {
            if(!this._endYear$p$0 && !this._endMonth$p$0 && !this._endDay$p$0)
                return true;
            var startDateTime = new Date;
            startDateTime["setFullYear"](this._startYear$p$0);
            startDateTime["setMonth"](this._startMonth$p$0 - 1);
            startDateTime["setDate"](this._startDay$p$0);
            var endDateTime = new Date;
            endDateTime["setFullYear"](this._endYear$p$0);
            endDateTime["setMonth"](this._endMonth$p$0 - 1);
            endDateTime["setDate"](this._endDay$p$0);
            return endDateTime >= startDateTime
        },
        _prependZeroToString$p$0: function(number)
        {
            if(number < 0)
                number = 1;
            if(number < 10)
                return"0" + number["toString"]();
            return number["toString"]()
        },
        _throwOnInvalidDateString$p$0: function(dateString)
        {
            var regEx = new RegExp("^\\d{4}-(?:[0]\\d|1[0-2])-(?:[0-2]\\d|3[01])$");
            if(!regEx["test"](dateString))
                throw Error.create(window["_u"]["ExtensibilityStrings"]["l_InvalidDate_Text"]);
        },
        _throwOnInvalidDate$p$0: function(year, month, day)
        {
            if(!this._isValidDate$p$0(year,month,day))
                throw Error.create(window["_u"]["ExtensibilityStrings"]["l_InvalidDate_Text"]);
        },
        _isValidDate$p$0: function(year, month, day)
        {
            if(year < 1601 || month < 1 || month > 12 || day < 1 || day > 31)
                return false;
            return true
        },
        _setDateHelper$p$0: function(isStart, yearOrDateString, month, day)
        {
            var yearCalculated = 0;
            var monthCalculated = 0;
            var dayCalculated = 0;
            if(yearOrDateString && !$h.ScriptHelpers.isNullOrUndefined(month) && day)
            {
                this._throwOnInvalidDate$p$0(yearOrDateString,month + 1,day);
                yearCalculated = yearOrDateString;
                monthCalculated = month + 1;
                dayCalculated = day
            }
            else if(yearOrDateString)
            {
                var dateString = yearOrDateString;
                this._throwOnInvalidDateString$p$0(dateString);
                var dateObject = new Date(dateString);
                if(dateObject && !window["isNaN"](dateObject["getUTCFullYear"]()) && !window["isNaN"](dateObject["getUTCMonth"]()) && !window["isNaN"](dateObject["getUTCDate"]()))
                {
                    this._throwOnInvalidDate$p$0(dateObject["getUTCFullYear"](),dateObject["getUTCMonth"]() + 1,dateObject["getUTCDate"]());
                    yearCalculated = dateObject["getUTCFullYear"]();
                    monthCalculated = dateObject["getUTCMonth"]() + 1;
                    dayCalculated = dateObject["getUTCDate"]()
                }
            }
            if(yearCalculated && monthCalculated && dayCalculated)
                if(isStart)
                {
                    this._startYear$p$0 = yearCalculated;
                    this._startMonth$p$0 = monthCalculated;
                    this._startDay$p$0 = dayCalculated
                }
                else
                {
                    this._endYear$p$0 = yearCalculated;
                    this._endMonth$p$0 = monthCalculated;
                    this._endDay$p$0 = dayCalculated
                }
        }
    };
    Microsoft.Office.WebExtension.SeriesTime.prototype.setStartDate = function(yearOrDateString, month, day)
    {
        if(yearOrDateString && !$h.ScriptHelpers.isNullOrUndefined(month) && day)
            this._setDateHelper$p$0(true,yearOrDateString,month,day);
        else if(yearOrDateString)
            this._setDateHelper$p$0(true,yearOrDateString,null,null)
    };
    Microsoft.Office.WebExtension.SeriesTime.prototype.getStartDate = function()
    {
        return this._startYear$p$0["toString"]() + "-" + this._prependZeroToString$p$0(this._startMonth$p$0) + "-" + this._prependZeroToString$p$0(this._startDay$p$0)
    };
    Microsoft.Office.WebExtension.SeriesTime.prototype.setEndDate = function(yearOrDateString, month, day)
    {
        if(yearOrDateString && !$h.ScriptHelpers.isNullOrUndefined(month) && day)
            this._setDateHelper$p$0(false,yearOrDateString,month,day);
        else if(yearOrDateString)
            this._setDateHelper$p$0(false,yearOrDateString,null,null);
        else if(!yearOrDateString)
        {
            this._endYear$p$0 = 0;
            this._endMonth$p$0 = 0;
            this._endDay$p$0 = 0
        }
    };
    Microsoft.Office.WebExtension.SeriesTime.prototype.getEndDate = function()
    {
        if(!this._endYear$p$0 && !this._endMonth$p$0 && !this._endDay$p$0)
            return null;
        return this._endYear$p$0["toString"]() + "-" + this._prependZeroToString$p$0(this._endMonth$p$0) + "-" + this._prependZeroToString$p$0(this._endDay$p$0)
    };
    Microsoft.Office.WebExtension.SeriesTime.prototype.setStartTime = function(hoursOrTimeString, minutes)
    {
        if(!$h.ScriptHelpers.isNullOrUndefined(hoursOrTimeString) && !$h.ScriptHelpers.isNullOrUndefined(minutes))
        {
            var totalMinutes = hoursOrTimeString * 60 + minutes;
            if(totalMinutes >= 0)
                this._startTimeMinutes$p$0 = totalMinutes;
            else
                throw Error.create(window["_u"]["ExtensibilityStrings"]["l_InvalidTime_Text"]);
        }
        else if(!$h.ScriptHelpers.isNullOrUndefined(hoursOrTimeString))
        {
            var timeString = hoursOrTimeString;
            var newDateString = "2017-01-15" + timeString + "Z";
            var RegEx = new RegExp("^T[0-2]\\d:[0-5]\\d:[0-5]\\d\\.\\d{3}$");
            if(!RegEx["test"](timeString))
                throw Error.create(window["_u"]["ExtensibilityStrings"]["l_InvalidTime_Text"]);
            var dateObject = new Date(newDateString);
            if(dateObject && !window["isNaN"](dateObject["getUTCHours"]()) && !window["isNaN"](dateObject["getUTCMinutes"]()))
                this._startTimeMinutes$p$0 = dateObject["getUTCHours"]() * 60 + dateObject["getUTCMinutes"]();
            else
                throw Error.create(window["_u"]["ExtensibilityStrings"]["l_InvalidTime_Text"]);
        }
    };
    Microsoft.Office.WebExtension.SeriesTime.prototype.getStartTime = function()
    {
        var minutes = this._startTimeMinutes$p$0 % 60;
        var hours = Math["floor"](this._startTimeMinutes$p$0 / 60);
        return"T" + this._prependZeroToString$p$0(hours) + ":" + this._prependZeroToString$p$0(minutes) + ":00.000"
    };
    Microsoft.Office.WebExtension.SeriesTime.prototype.getEndTime = function()
    {
        var endTimeMinutes = this._startTimeMinutes$p$0 + this._durationMinutes$p$0;
        var minutes = endTimeMinutes % 60;
        var hours = Math["floor"](endTimeMinutes / 60) % 24;
        return"T" + this._prependZeroToString$p$0(hours) + ":" + this._prependZeroToString$p$0(minutes) + ":00.000"
    };
    Microsoft.Office.WebExtension.SeriesTime.prototype.setDuration = function(minutes)
    {
        if(minutes >= 0)
            this._durationMinutes$p$0 = minutes;
        else
            throw Error.create(window["_u"]["ExtensibilityStrings"]["l_InvalidTime_Text"]);
    };
    Microsoft.Office.WebExtension.SeriesTime.prototype.getDuration = function()
    {
        return this._durationMinutes$p$0
    };
    $h.ReplyConstants = function(){};
    $h.EmailAddressConstants = function(){};
    $h.CategoriesConstants = function(){};
    $h.AsyncConstants = function(){};
    $h.ApiTelemetryCode = function(){};
    window["Office"]["cast"]["item"] = Office.cast.item = function(){};
    window["Office"]["cast"]["item"]["toItemRead"] = function(item)
    {
        if($h.Item["isInstanceOfType"](item))
            return item;
        throw Error.argumentType();
    };
    window["Office"]["cast"]["item"]["toItemCompose"] = function(item)
    {
        if($h.ComposeItem["isInstanceOfType"](item))
            return item;
        throw Error.argumentType();
    };
    window["Office"]["cast"]["item"]["toMessage"] = function(item)
    {
        return window["Office"]["cast"]["item"]["toMessageRead"](item)
    };
    window["Office"]["cast"]["item"]["toMessageRead"] = function(item)
    {
        if($h.Message["isInstanceOfType"](item))
            return item;
        throw Error.argumentType();
    };
    window["Office"]["cast"]["item"]["toMessageCompose"] = function(item)
    {
        if($h.MessageCompose["isInstanceOfType"](item))
            return item;
        throw Error.argumentType();
    };
    window["Office"]["cast"]["item"]["toMeetingRequest"] = function(item)
    {
        if($h.MeetingRequest["isInstanceOfType"](item))
            return item;
        throw Error.argumentType();
    };
    window["Office"]["cast"]["item"]["toAppointment"] = function(item)
    {
        return window["Office"]["cast"]["item"]["toAppointmentRead"](item)
    };
    window["Office"]["cast"]["item"]["toAppointmentRead"] = function(item)
    {
        if($h.Appointment["isInstanceOfType"](item))
            return item;
        throw Error.argumentType();
    };
    window["Office"]["cast"]["item"]["toAppointmentCompose"] = function(item)
    {
        if($h.AppointmentCompose["isInstanceOfType"](item))
            return item;
        throw Error.argumentType();
    };
    $h.Item = function(data)
    {
        this.$$d__getBody$p$1 = Function.createDelegate(this,this._getBody$p$1);
        this.$$d__getAttachments$p$1 = Function.createDelegate(this,this._getAttachments$p$1);
        this.$$d__getItemClass$p$1 = Function.createDelegate(this,this._getItemClass$p$1);
        this.$$d__getItemId$p$1 = Function.createDelegate(this,this._getItemId$p$1);
        this.$$d__getDateTimeModified$p$1 = Function.createDelegate(this,this._getDateTimeModified$p$1);
        this.$$d__getDateTimeCreated$p$1 = Function.createDelegate(this,this._getDateTimeCreated$p$1);
        $h.Item["initializeBase"](this,[data]);
        $h.InitialData._defineReadOnlyProperty$i(this,"dateTimeCreated",this.$$d__getDateTimeCreated$p$1);
        $h.InitialData._defineReadOnlyProperty$i(this,"dateTimeModified",this.$$d__getDateTimeModified$p$1);
        $h.InitialData._defineReadOnlyProperty$i(this,"itemId",this.$$d__getItemId$p$1);
        $h.InitialData._defineReadOnlyProperty$i(this,"itemClass",this.$$d__getItemClass$p$1);
        $h.InitialData._defineReadOnlyProperty$i(this,"attachments",this.$$d__getAttachments$p$1);
        $h.InitialData._defineReadOnlyProperty$i(this,"body",this.$$d__getBody$p$1)
    };
    $h.Item.prototype = {
        _body$p$1: null,
        _getItemId$p$1: function()
        {
            return this._data$p$0.get__itemId$i$0()
        },
        _getItemClass$p$1: function()
        {
            return this._data$p$0.get__itemClass$i$0()
        },
        _getDateTimeCreated$p$1: function()
        {
            return this._data$p$0.get__dateTimeCreated$i$0()
        },
        _getDateTimeModified$p$1: function()
        {
            return this._data$p$0.get__dateTimeModified$i$0()
        },
        _getAttachments$p$1: function()
        {
            return this._data$p$0.get__attachments$i$0()
        },
        _getBody$p$1: function()
        {
            if(!this._body$p$1)
                this._body$p$1 = new $h.Body;
            return this._body$p$1
        },
        _validateDestinationFolder$p$1: function(destinationFolder)
        {
            if(destinationFolder !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Folder"]["Inbox"] && destinationFolder !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Folder"]["Junk"] && destinationFolder !== window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["Folder"]["DeletedItems"])
                throw Error.argument("destinationFolder");
        }
    };
    $h.Item.prototype.move = function(destinationFolder)
    {
        var args = [];
        for(var $$pai_4 = 1; $$pai_4 < arguments["length"]; ++$$pai_4)
            args[$$pai_4 - 1] = arguments[$$pai_4];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(3,"item.move");
        this._validateDestinationFolder$p$1(destinationFolder);
        var commonParameters = $h.CommonParameters.parse(args,false);
        var dataToHost = {destinationFolder: destinationFolder};
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(101,dataToHost,null,null,commonParameters._callback$p$0)
    };
    $h.ItemBase = function(data)
    {
        this.$$d__createCustomProperties$i$0 = Function.createDelegate(this,this._createCustomProperties$i$0);
        this.$$d__getCategories$p$0 = Function.createDelegate(this,this._getCategories$p$0);
        this.$$d__getNotificationMessages$p$0 = Function.createDelegate(this,this._getNotificationMessages$p$0);
        this.$$d_getItemType = Function.createDelegate(this,this.getItemType);
        this._data$p$0 = data;
        $h.InitialData._defineReadOnlyProperty$i(this,"itemType",this.$$d_getItemType);
        $h.InitialData._defineReadOnlyProperty$i(this,"notificationMessages",this.$$d__getNotificationMessages$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"categories",this.$$d__getCategories$p$0)
    };
    $h.ItemBase.prototype = {
        _categories$p$0: null,
        get__isFromSharedFolder$i$0: function()
        {
            return this._data$p$0.get__isFromSharedFolder$i$0()
        },
        _data$p$0: null,
        _notificationMessages$p$0: null,
        get_data: function()
        {
            return this._data$p$0
        },
        _createCustomProperties$i$0: function(data)
        {
            return new $h.CustomProperties(data)
        },
        _getNotificationMessages$p$0: function()
        {
            if(!this._notificationMessages$p$0)
                this._notificationMessages$p$0 = new $h.NotificationMessages;
            return this._notificationMessages$p$0
        },
        _getCategories$p$0: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._categories$p$0)
                this._categories$p$0 = new $h.Categories;
            return this._categories$p$0
        }
    };
    $h.ItemBase.prototype.loadCustomPropertiesAsync = function()
    {
        var args = [];
        for(var $$pai_3 = 0; $$pai_3 < arguments["length"]; ++$$pai_3)
            args[$$pai_3] = arguments[$$pai_3];
        var parameters = $h.CommonParameters.parse(args,true,true);
        var loadCustomProperties = new $h._loadDictionaryRequest(this.$$d__createCustomProperties$i$0,"customProperties",parameters._callback$p$0,parameters._asyncContext$p$0);
        loadCustomProperties._sendRequest$i$0(3,"LoadCustomProperties",{})
    };
    $h.ItemBase.prototype.getInitializationContextAsync = function()
    {
        var args = [];
        for(var $$pai_2 = 0; $$pai_2 < arguments["length"]; ++$$pai_2)
            args[$$pai_2] = arguments[$$pai_2];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"getInitializationContextAsync");
        var parameters = $h.CommonParameters.parse(args,true);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(99,null,null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.ItemBase.prototype.getAttachmentContentAsync = function(attachmentId)
    {
        var args = [];
        for(var $$pai_3 = 1; $$pai_3 < arguments["length"]; ++$$pai_3)
            args[$$pai_3 - 1] = arguments[$$pai_3];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"getAttachmentContentAsync");
        if(!$h.ScriptHelpers.isNonEmptyString(attachmentId))
            throw Error.argument("attachmentId");
        var commonParameters = $h.CommonParameters.parse(args,true);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(150,{id: attachmentId},null,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
    };
    $h.MasterCategories = function(){};
    $h.MasterCategories.prototype = {
        _validateCategoryDetailsArray$p$0: function(categoryDetails)
        {
            if($h.ScriptHelpers.isNullOrUndefined(categoryDetails))
                throw Error.argument("categoryDetails");
            if(!Array["isInstanceOfType"](categoryDetails))
                throw Error.argumentType("categoryDetails",Object["getType"](categoryDetails),Array);
            if(!categoryDetails["length"])
                throw Error.argument("categoryDetails");
            for(var $$arr_1 = categoryDetails, $$len_2 = $$arr_1.length, $$idx_3 = 0; $$idx_3 < $$len_2; ++$$idx_3)
            {
                var categoryDetailsDictionary = $$arr_1[$$idx_3];
                this._validateCategoryDetailsDictionary$p$0(categoryDetailsDictionary)
            }
        },
        _validateCategoryDetailsDictionary$p$0: function(dict)
        {
            if($h.ScriptHelpers.isNullOrUndefined(dict))
                throw Error.argument("categoryDetails");
            var keys = Object["keys"](dict);
            if(keys["length"] !== 2)
                throw Error.argument("categoryDetails");
            var displayName = dict["displayName"];
            this._validateCategoryDetailsDisplayNameParameter$p$0(displayName);
            var color = dict["color"];
            this._validateCategoryDetailsColorParameter$p$0(color)
        },
        _validateCategoryDetailsDisplayNameParameter$p$0: function(displayName)
        {
            if(!$h.ScriptHelpers.isNonEmptyString(displayName) || displayName.length > 255)
                throw Error.argument("displayName");
        },
        _validateCategoryDetailsColorParameter$p$0: function(color)
        {
            if(!$h.ScriptHelpers.isNonEmptyString(color) || Array.indexOf($h.MasterCategories._colorPresets$i,color) === -1)
                throw Error.argument("color");
        }
    };
    $h.MasterCategories.prototype.getAsync = function()
    {
        var args = [];
        for(var $$pai_2 = 0; $$pai_2 < arguments["length"]; ++$$pai_2)
            args[$$pai_2] = arguments[$$pai_2];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(3,"masterCategories.getAsync");
        var parameters = $h.CommonParameters.parse(args,true);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(160,null,null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.MasterCategories.prototype.addAsync = function(categoryDetailsArray)
    {
        var args = [];
        for(var $$pai_3 = 1; $$pai_3 < arguments["length"]; ++$$pai_3)
            args[$$pai_3 - 1] = arguments[$$pai_3];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(3,"masterCategories.addAsync");
        var parameters = $h.CommonParameters.parse(args,true);
        this._validateCategoryDetailsArray$p$0(categoryDetailsArray);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(161,{categoryDetails: categoryDetailsArray},null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.MasterCategories.prototype.removeAsync = function(categories)
    {
        var args = [];
        for(var $$pai_3 = 1; $$pai_3 < arguments["length"]; ++$$pai_3)
            args[$$pai_3 - 1] = arguments[$$pai_3];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(3,"masterCategories.removeAsync");
        var parameters = $h.CommonParameters.parse(args,false);
        $h.ScriptHelpers.validateCategoriesArray(categories);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(162,{categories: categories},null,parameters._asyncContext$p$0,parameters._callback$p$0)
    };
    $h.MeetingRequest = function(data)
    {
        this.$$d__getEnhancedLocation$p$3 = Function.createDelegate(this,this._getEnhancedLocation$p$3);
        this.$$d__getSeriesId$p$3 = Function.createDelegate(this,this._getSeriesId$p$3);
        this.$$d__getRecurrence$p$3 = Function.createDelegate(this,this._getRecurrence$p$3);
        this.$$d__getRequiredAttendees$p$3 = Function.createDelegate(this,this._getRequiredAttendees$p$3);
        this.$$d__getOptionalAttendees$p$3 = Function.createDelegate(this,this._getOptionalAttendees$p$3);
        this.$$d__getLocation$p$3 = Function.createDelegate(this,this._getLocation$p$3);
        this.$$d__getEnd$p$3 = Function.createDelegate(this,this._getEnd$p$3);
        this.$$d__getStart$p$3 = Function.createDelegate(this,this._getStart$p$3);
        $h.MeetingRequest["initializeBase"](this,[data]);
        $h.InitialData._defineReadOnlyProperty$i(this,"start",this.$$d__getStart$p$3);
        $h.InitialData._defineReadOnlyProperty$i(this,"end",this.$$d__getEnd$p$3);
        $h.InitialData._defineReadOnlyProperty$i(this,"location",this.$$d__getLocation$p$3);
        $h.InitialData._defineReadOnlyProperty$i(this,"optionalAttendees",this.$$d__getOptionalAttendees$p$3);
        $h.InitialData._defineReadOnlyProperty$i(this,"requiredAttendees",this.$$d__getRequiredAttendees$p$3);
        $h.InitialData._defineReadOnlyProperty$i(this,"recurrence",this.$$d__getRecurrence$p$3);
        $h.InitialData._defineReadOnlyProperty$i(this,"seriesId",this.$$d__getSeriesId$p$3);
        $h.InitialData._defineReadOnlyProperty$i(this,"enhancedLocation",this.$$d__getEnhancedLocation$p$3)
    };
    $h.MeetingRequest.prototype = {
        _enhancedLocation$p$3: null,
        _getStart$p$3: function()
        {
            return this._data$p$0.get__start$i$0()
        },
        _getEnd$p$3: function()
        {
            return this._data$p$0.get__end$i$0()
        },
        _getLocation$p$3: function()
        {
            return this._data$p$0.get__location$i$0()
        },
        _getOptionalAttendees$p$3: function()
        {
            return this._data$p$0.get__cc$i$0()
        },
        _getRequiredAttendees$p$3: function()
        {
            return this._data$p$0.get__to$i$0()
        },
        _getRecurrence$p$3: function()
        {
            if(this._data$p$0.get__recurrence$i$0() && this._data$p$0.get__recurrence$i$0()["seriesTimeJson"])
                return $h.ComposeRecurrence.copyRecurrenceObjectConvertSeriesTimeJson(this._data$p$0.get__recurrence$i$0());
            return this._data$p$0.get__recurrence$i$0()
        },
        _getSeriesId$p$3: function()
        {
            return this._data$p$0.get__seriesId$i$0()
        },
        _getEnhancedLocation$p$3: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._enhancedLocation$p$3)
                this._enhancedLocation$p$3 = new $h.EnhancedLocation(false);
            return this._enhancedLocation$p$3
        }
    };
    $h.MeetingSuggestion = function(data, dateTimeSent)
    {
        this.$$d__getEndTime$p$0 = Function.createDelegate(this,this._getEndTime$p$0);
        this.$$d__getStartTime$p$0 = Function.createDelegate(this,this._getStartTime$p$0);
        this.$$d__getSubject$p$0 = Function.createDelegate(this,this._getSubject$p$0);
        this.$$d__getLocation$p$0 = Function.createDelegate(this,this._getLocation$p$0);
        this.$$d__getAttendees$p$0 = Function.createDelegate(this,this._getAttendees$p$0);
        this.$$d__getMeetingString$p$0 = Function.createDelegate(this,this._getMeetingString$p$0);
        this._data$p$0 = data;
        this._dateTimeSent$p$0 = dateTimeSent;
        $h.InitialData._defineReadOnlyProperty$i(this,"meetingString",this.$$d__getMeetingString$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"attendees",this.$$d__getAttendees$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"location",this.$$d__getLocation$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"subject",this.$$d__getSubject$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"start",this.$$d__getStartTime$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"end",this.$$d__getEndTime$p$0)
    };
    $h.MeetingSuggestion.prototype = {
        _dateTimeSent$p$0: null,
        _data$p$0: null,
        _attendees$p$0: null,
        _getMeetingString$p$0: function()
        {
            return this._data$p$0["MeetingString"]
        },
        _getLocation$p$0: function()
        {
            return this._data$p$0["Location"]
        },
        _getSubject$p$0: function()
        {
            return this._data$p$0["Subject"]
        },
        _isUTC$p$0: function()
        {
            if(!("IsLegacyEntityExtraction" in this._data$p$0))
                return true;
            return this._data$p$0["IsLegacyEntityExtraction"]
        },
        _getStartTime$p$0: function()
        {
            var time = this._createDateTimeFromParameter$p$0("StartTime");
            var resolvedTime = $h.MeetingSuggestionTimeDecoder.resolve(time,this._dateTimeSent$p$0,this._isUTC$p$0());
            if(resolvedTime["getTime"]() !== time["getTime"]())
                return window["OSF"]["DDA"]["OutlookAppOm"]._instance$p["convertToUtcClientTime"](window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._dateToDictionary$i$0(resolvedTime));
            return time
        },
        _getEndTime$p$0: function()
        {
            var time = this._createDateTimeFromParameter$p$0("EndTime");
            var resolvedTime = $h.MeetingSuggestionTimeDecoder.resolve(time,this._dateTimeSent$p$0,this._isUTC$p$0());
            if(resolvedTime["getTime"]() !== time["getTime"]())
                return window["OSF"]["DDA"]["OutlookAppOm"]._instance$p["convertToUtcClientTime"](window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._dateToDictionary$i$0(resolvedTime));
            return time
        },
        _createDateTimeFromParameter$p$0: function(keyName)
        {
            var dateTimeString = this._data$p$0[keyName];
            if(!dateTimeString)
                return null;
            return new Date(dateTimeString)
        },
        _getAttendees$p$0: function()
        {
            if(!this._attendees$p$0)
            {
                var $$t_1 = this;
                this._attendees$p$0 = $h.Entities._getExtractedObjects$i($h.EmailAddressDetails,this._data$p$0,"Attendees",function(data)
                {
                    return $h.EmailAddressDetails._createFromEmailUserDictionary$i(data)
                })
            }
            return this._attendees$p$0
        }
    };
    $h.MeetingSuggestionTimeDecoder = function(){};
    $h.MeetingSuggestionTimeDecoder.resolve = function(inTime, sentTime, isUTC)
    {
        if(!sentTime)
            return inTime;
        try
        {
            var tod;
            var outDate;
            var extractedDate;
            var sentDate = new Date(sentTime["getFullYear"](),sentTime["getMonth"](),sentTime["getDate"](),0,0,0,0);
            var $$t_8,
                $$t_9,
                $$t_A;
            if(!($$t_A = $h.MeetingSuggestionTimeDecoder._decode$p(inTime,isUTC,$$t_8 = {val: extractedDate},$$t_9 = {val: tod}),extractedDate = $$t_8["val"],tod = $$t_9["val"],$$t_A))
                return inTime;
            else
            {
                if($h._preciseDate["isInstanceOfType"](extractedDate))
                    outDate = $h.MeetingSuggestionTimeDecoder._resolvePreciseDate$p(sentDate,extractedDate);
                else if($h._relativeDate["isInstanceOfType"](extractedDate))
                    outDate = $h.MeetingSuggestionTimeDecoder._resolveRelativeDate$p(sentDate,extractedDate);
                else
                    outDate = sentDate;
                if(window["isNaN"](outDate["getTime"]()))
                    return sentTime;
                outDate["setMilliseconds"](outDate["getMilliseconds"]() + tod);
                return outDate
            }
        }
        catch($$e_7)
        {
            return sentTime
        }
    };
    $h.MeetingSuggestionTimeDecoder._isNullOrUndefined$i = function(value)
    {
        return null === value || value === undefined
    };
    $h.MeetingSuggestionTimeDecoder._resolvePreciseDate$p = function(sentDate, precise)
    {
        var year = precise._year$i$1;
        var month = !precise._month$i$1 ? sentDate["getMonth"]() : precise._month$i$1 - 1;
        var day = precise._day$i$1;
        if(!day)
            return sentDate;
        var candidate;
        if($h.MeetingSuggestionTimeDecoder._isNullOrUndefined$i(year))
        {
            candidate = new Date(sentDate["getFullYear"](),month,day);
            if(candidate["getTime"]() < sentDate["getTime"]())
                candidate = new Date(sentDate["getFullYear"]() + 1,month,day)
        }
        else
            candidate = new Date(year < 50 ? 2e3 + year : 1900 + year,month,day);
        if(candidate["getMonth"]() !== month)
            return sentDate;
        return candidate
    };
    $h.MeetingSuggestionTimeDecoder._resolveRelativeDate$p = function(sentDate, relative)
    {
        var date;
        switch(relative._unit$i$1)
        {
            case 0:
                date = new Date(sentDate["getFullYear"](),sentDate["getMonth"](),sentDate["getDate"]());
                date["setDate"](date["getDate"]() + relative._offset$i$1);
                return date;
            case 5:
                return $h.MeetingSuggestionTimeDecoder._findBestDateForWeekDate$p(sentDate,relative._offset$i$1,relative._tag$i$1);
            case 2:
                var days = 1;
                switch(relative._modifier$i$1)
                {
                    case 1:
                        break;
                    case 2:
                        days = 16;
                        break;
                    default:
                        if(!relative._offset$i$1)
                            days = sentDate["getDate"]();
                        break
                }
                date = new Date(sentDate["getFullYear"](),sentDate["getMonth"](),days);
                date["setMonth"](date["getMonth"]() + relative._offset$i$1);
                if(date["getTime"]() < sentDate["getTime"]())
                    date["setDate"](date["getDate"]() + sentDate["getDate"]() - 1);
                return date;
            case 1:
                date = new Date(sentDate["getFullYear"](),sentDate["getMonth"](),sentDate["getDate"]());
                date["setDate"](sentDate["getDate"]() + 7 * relative._offset$i$1);
                if(relative._modifier$i$1 === 1 || !relative._modifier$i$1)
                {
                    date["setDate"](date["getDate"]() + 1 - date["getDay"]());
                    if(date["getTime"]() < sentDate["getTime"]())
                        return sentDate;
                    return date
                }
                else if(relative._modifier$i$1 === 2)
                {
                    date["setDate"](date["getDate"]() + 5 - date["getDay"]());
                    return date
                }
                break;
            case 4:
                return $h.MeetingSuggestionTimeDecoder._findBestDateForWeekOfMonthDate$p(sentDate,relative);
            case 3:
                if(relative._offset$i$1 > 0)
                    return new Date(sentDate["getFullYear"]() + relative._offset$i$1,0,1);
                break;
            default:
                break
        }
        return sentDate
    };
    $h.MeetingSuggestionTimeDecoder._findBestDateForWeekDate$p = function(sentDate, offset, tag)
    {
        if(offset > -5 && offset < 5)
        {
            var dayOfWeek = (tag + 6) % 7 + 1;
            var days = 7 * offset + (dayOfWeek - sentDate["getDay"]());
            sentDate["setDate"](sentDate["getDate"]() + days);
            return sentDate
        }
        else
        {
            var days = (tag - sentDate["getDay"]()) % 7;
            if(days < 0)
                days += 7;
            sentDate["setDate"](sentDate["getDate"]() + days);
            return sentDate
        }
    };
    $h.MeetingSuggestionTimeDecoder._findBestDateForWeekOfMonthDate$p = function(sentDate, relative)
    {
        var date;
        var firstDay;
        var newDate;
        date = sentDate;
        if(relative._tag$i$1 <= 0 || relative._tag$i$1 > 12 || relative._offset$i$1 <= 0 || relative._offset$i$1 > 5)
            return sentDate;
        var monthOffset = (12 + relative._tag$i$1 - date["getMonth"]() - 1) % 12;
        firstDay = new Date(date["getFullYear"](),date["getMonth"]() + monthOffset,1);
        if(relative._modifier$i$1 === 1)
            if(relative._offset$i$1 === 1 && firstDay["getDay"]() !== 6 && firstDay["getDay"]())
                return firstDay;
            else
            {
                newDate = new Date(firstDay["getFullYear"](),firstDay["getMonth"](),firstDay["getDate"]());
                newDate["setDate"](newDate["getDate"]() + (7 + (1 - firstDay["getDay"]())) % 7);
                if(firstDay["getDay"]() !== 6 && firstDay["getDay"]() && firstDay["getDay"]() !== 1)
                    newDate["setDate"](newDate["getDate"]() - 7);
                newDate["setDate"](newDate["getDate"]() + 7 * (relative._offset$i$1 - 1));
                if(newDate["getMonth"]() + 1 !== relative._tag$i$1)
                    return sentDate;
                return newDate
            }
        else
        {
            newDate = new Date(firstDay["getFullYear"](),firstDay["getMonth"](),$h.MeetingSuggestionTimeDecoder._daysInMonth$p(firstDay["getMonth"](),firstDay["getFullYear"]()));
            var offset = 1 - newDate["getDay"]();
            if(offset > 0)
                offset = offset - 7;
            newDate["setDate"](newDate["getDate"]() + offset);
            newDate["setDate"](newDate["getDate"]() + 7 * (1 - relative._offset$i$1));
            if(newDate["getMonth"]() + 1 !== relative._tag$i$1)
                if(firstDay["getDay"]() !== 6 && firstDay["getDay"]())
                    return firstDay;
                else
                    return sentDate;
            else
                return newDate
        }
    };
    $h.MeetingSuggestionTimeDecoder._decode$p = function(inDate, isUTC, date, time)
    {
        var DateValueMask = 32767;
        date["val"] = null;
        time["val"] = 0;
        if(!inDate)
            return false;
        if(isUTC)
            time["val"] = $h.MeetingSuggestionTimeDecoder._getTimeOfDayInMillisecondsUTC$p(inDate);
        else
            time["val"] = $h.MeetingSuggestionTimeDecoder._getTimeOfDayInMilliseconds$p(inDate);
        var inDateAtMidnight = inDate["getTime"]() - time["val"];
        var value = (inDateAtMidnight - $h.MeetingSuggestionTimeDecoder._baseDate$p["getTime"]()) / 864e5;
        if(value < 0)
            return false;
        else if(value >= 262144)
            return false;
        else
        {
            var type = value >> 15;
            value = value & DateValueMask;
            switch(type)
            {
                case 0:
                    return $h.MeetingSuggestionTimeDecoder._decodePreciseDate$p(value,date);
                case 1:
                    return $h.MeetingSuggestionTimeDecoder._decodeRelativeDate$p(value,date);
                default:
                    return false
            }
        }
    };
    $h.MeetingSuggestionTimeDecoder._decodePreciseDate$p = function(value, date)
    {
        var c_SubTypeMask = 7;
        var c_MonthMask = 15;
        var c_DayMask = 31;
        var c_YearMask = 127;
        var year = null;
        var month = 0;
        var day = 0;
        date["val"] = null;
        var subType = value >> 12 & c_SubTypeMask;
        if((subType & 4) === 4)
        {
            year = value >> 5 & c_YearMask;
            if((subType & 2) === 2)
            {
                if((subType & 1) === 1)
                    return false;
                month = value >> 1 & c_MonthMask
            }
        }
        else
        {
            if((subType & 2) === 2)
                month = value >> 8 & c_MonthMask;
            if((subType & 1) === 1)
                day = value >> 3 & c_DayMask
        }
        date["val"] = new $h._preciseDate(day,month,year);
        return true
    };
    $h.MeetingSuggestionTimeDecoder._decodeRelativeDate$p = function(value, date)
    {
        var TagMask = 15;
        var OffsetMask = 63;
        var UnitMask = 7;
        var ModifierMask = 3;
        var tag = value & TagMask;
        value >>= 4;
        var offset = $h.MeetingSuggestionTimeDecoder._fromComplement$p(value & OffsetMask,6);
        value >>= 6;
        var unit = value & UnitMask;
        value >>= 3;
        var modifier = value & ModifierMask;
        try
        {
            date["val"] = new $h._relativeDate(modifier,offset,unit,tag);
            return true
        }
        catch($$e_A)
        {
            date["val"] = null;
            return false
        }
    };
    $h.MeetingSuggestionTimeDecoder._fromComplement$p = function(value, n)
    {
        var signed = 1 << n - 1;
        var mask = (1 << n) - 1;
        if((value & signed) === signed)
            return-((value ^ mask) + 1);
        else
            return value
    };
    $h.MeetingSuggestionTimeDecoder._daysInMonth$p = function(month, year)
    {
        return 32 - new Date(year,month,32)["getDate"]()
    };
    $h.MeetingSuggestionTimeDecoder._getTimeOfDayInMilliseconds$p = function(inputTime)
    {
        var timeOfDay = 0;
        timeOfDay += inputTime["getHours"]() * 3600;
        timeOfDay += inputTime["getMinutes"]() * 60;
        timeOfDay += inputTime["getSeconds"]();
        timeOfDay *= 1e3;
        timeOfDay += inputTime["getMilliseconds"]();
        return timeOfDay
    };
    $h.MeetingSuggestionTimeDecoder._getTimeOfDayInMillisecondsUTC$p = function(inputTime)
    {
        var timeOfDay = 0;
        timeOfDay += inputTime["getUTCHours"]() * 3600;
        timeOfDay += inputTime["getUTCMinutes"]() * 60;
        timeOfDay += inputTime["getUTCSeconds"]();
        timeOfDay *= 1e3;
        timeOfDay += inputTime["getUTCMilliseconds"]();
        return timeOfDay
    };
    $h._extractedDate = function(){};
    $h._preciseDate = function(day, month, year)
    {
        $h._preciseDate["initializeBase"](this);
        if(day < 0 || day > 31)
            throw Error.argumentOutOfRange("day");
        if(month < 0 || month > 12)
            throw Error.argumentOutOfRange("month");
        this._day$i$1 = day;
        this._month$i$1 = month;
        if(!$h.MeetingSuggestionTimeDecoder._isNullOrUndefined$i(year))
        {
            if(!month && day)
                throw Error.argument("Invalid arguments");
            if(year < 0 || year > 2099)
                throw Error.argumentOutOfRange("year");
            this._year$i$1 = year % 100
        }
        else if(!this._month$i$1 && !this._day$i$1)
            throw Error.argument("Invalid datetime");
    };
    $h._preciseDate.prototype = {
        _day$i$1: 0,
        _month$i$1: 0,
        _year$i$1: null
    };
    $h._relativeDate = function(modifier, offset, unit, tag)
    {
        $h._relativeDate["initializeBase"](this);
        if(offset < -32 || offset > 31)
            throw Error.argumentOutOfRange("offset");
        if(tag < 0 || tag > 15)
            throw Error.argumentOutOfRange("tag");
        if(!unit && offset < 0)
            throw Error.argument("unit & offset do not form a valid date");
        this._modifier$i$1 = modifier;
        this._offset$i$1 = offset;
        this._unit$i$1 = unit;
        this._tag$i$1 = tag
    };
    $h._relativeDate.prototype = {
        _modifier$i$1: 0,
        _offset$i$1: 0,
        _unit$i$1: 0,
        _tag$i$1: 0
    };
    $h.Message = function(dataDictionary)
    {
        this.$$d__getInternetHeaders$p$2 = Function.createDelegate(this,this._getInternetHeaders$p$2);
        this.$$d__getConversationId$p$2 = Function.createDelegate(this,this._getConversationId$p$2);
        this.$$d__getInternetMessageId$p$2 = Function.createDelegate(this,this._getInternetMessageId$p$2);
        this.$$d__getCc$p$2 = Function.createDelegate(this,this._getCc$p$2);
        this.$$d__getTo$p$2 = Function.createDelegate(this,this._getTo$p$2);
        this.$$d__getFrom$p$2 = Function.createDelegate(this,this._getFrom$p$2);
        this.$$d__getSender$p$2 = Function.createDelegate(this,this._getSender$p$2);
        this.$$d__getNormalizedSubject$p$2 = Function.createDelegate(this,this._getNormalizedSubject$p$2);
        this.$$d__getSubject$p$2 = Function.createDelegate(this,this._getSubject$p$2);
        $h.Message["initializeBase"](this,[dataDictionary]);
        $h.InitialData._defineReadOnlyProperty$i(this,"subject",this.$$d__getSubject$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"normalizedSubject",this.$$d__getNormalizedSubject$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"sender",this.$$d__getSender$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"from",this.$$d__getFrom$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"to",this.$$d__getTo$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"cc",this.$$d__getCc$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"internetMessageId",this.$$d__getInternetMessageId$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"conversationId",this.$$d__getConversationId$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"internetHeaders",this.$$d__getInternetHeaders$p$2)
    };
    $h.Message.prototype = {
        _internetHeaders$p$2: null,
        getItemType: function()
        {
            return window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ItemType"]["Message"]
        },
        _getSubject$p$2: function()
        {
            return this._data$p$0.get__subject$i$0()
        },
        _getNormalizedSubject$p$2: function()
        {
            return this._data$p$0.get__normalizedSubject$i$0()
        },
        _getSender$p$2: function()
        {
            return this._data$p$0.get__sender$i$0()
        },
        _getFrom$p$2: function()
        {
            return this._data$p$0.get__from$i$0()
        },
        _getTo$p$2: function()
        {
            return this._data$p$0.get__to$i$0()
        },
        _getCc$p$2: function()
        {
            return this._data$p$0.get__cc$i$0()
        },
        _getInternetMessageId$p$2: function()
        {
            return this._data$p$0.get__internetMessageId$i$0()
        },
        _getConversationId$p$2: function()
        {
            return this._data$p$0.get__conversationId$i$0()
        },
        _getInternetHeaders$p$2: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._internetHeaders$p$2)
                this._internetHeaders$p$2 = new $h.InternetHeaders(false);
            return this._internetHeaders$p$2
        }
    };
    $h.Message.prototype.getEntities = function()
    {
        return this._data$p$0._getEntities$i$0()
    };
    $h.Message.prototype.getEntitiesByType = function(entityType)
    {
        return this._data$p$0._getEntitiesByType$i$0(entityType)
    };
    $h.Message.prototype.getFilteredEntitiesByName = function(name)
    {
        return this._data$p$0._getFilteredEntitiesByName$i$0(name)
    };
    $h.Message.prototype.getSelectedEntities = function()
    {
        return this._data$p$0._getSelectedEntities$i$0()
    };
    $h.Message.prototype.getRegExMatches = function()
    {
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"getRegExMatches");
        return this._data$p$0._getRegExMatches$i$0()
    };
    $h.Message.prototype.getRegExMatchesByName = function(name)
    {
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"getRegExMatchesByName");
        return this._data$p$0._getRegExMatchesByName$i$0(name)
    };
    $h.Message.prototype.getSelectedRegExMatches = function()
    {
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(1,"getSelectedRegExMatches");
        return this._data$p$0._getSelectedRegExMatches$i$0()
    };
    $h.Message.prototype.displayReplyForm = function(obj)
    {
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._displayReplyForm$i$0(obj)
    };
    $h.Message.prototype.displayReplyAllForm = function(obj)
    {
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._displayReplyAllForm$i$0(obj)
    };
    $h.MessageCompose = function(data)
    {
        this.$$d__getInternetHeaders$p$2 = Function.createDelegate(this,this._getInternetHeaders$p$2);
        this.$$d__getFrom$p$2 = Function.createDelegate(this,this._getFrom$p$2);
        this.$$d__getConversationId$p$2 = Function.createDelegate(this,this._getConversationId$p$2);
        this.$$d__getBcc$p$2 = Function.createDelegate(this,this._getBcc$p$2);
        this.$$d__getCc$p$2 = Function.createDelegate(this,this._getCc$p$2);
        this.$$d__getTo$p$2 = Function.createDelegate(this,this._getTo$p$2);
        $h.MessageCompose["initializeBase"](this,[data]);
        $h.InitialData._defineReadOnlyProperty$i(this,"to",this.$$d__getTo$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"cc",this.$$d__getCc$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"bcc",this.$$d__getBcc$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"conversationId",this.$$d__getConversationId$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"from",this.$$d__getFrom$p$2);
        $h.InitialData._defineReadOnlyProperty$i(this,"internetHeaders",this.$$d__getInternetHeaders$p$2)
    };
    $h.MessageCompose.prototype = {
        _from$p$2: null,
        _to$p$2: null,
        _cc$p$2: null,
        _bcc$p$2: null,
        _internetHeaders$p$2: null,
        getItemType: function()
        {
            return window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ItemType"]["Message"]
        },
        _getTo$p$2: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._to$p$2)
                this._to$p$2 = new $h.ComposeRecipient(0,"to");
            return this._to$p$2
        },
        _getCc$p$2: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._cc$p$2)
                this._cc$p$2 = new $h.ComposeRecipient(1,"cc");
            return this._cc$p$2
        },
        _getBcc$p$2: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._bcc$p$2)
                this._bcc$p$2 = new $h.ComposeRecipient(2,"bcc");
            return this._bcc$p$2
        },
        _getConversationId$p$2: function()
        {
            return this._data$p$0.get__conversationId$i$0()
        },
        _getFrom$p$2: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._from$p$2)
                this._from$p$2 = new $h.ComposeFrom;
            return this._from$p$2
        },
        _getInternetHeaders$p$2: function()
        {
            this._data$p$0._throwOnRestrictedPermissionLevel$i$0();
            if(!this._internetHeaders$p$2)
                this._internetHeaders$p$2 = new $h.InternetHeaders(true);
            return this._internetHeaders$p$2
        }
    };
    $h.NotificationMessages = function(){};
    $h.NotificationMessages._mapToHostItemNotificationMessageType$p = function(dataToHost)
    {
        var notificationType;
        var hostItemNotificationMessageType;
        notificationType = dataToHost["type"];
        if(notificationType === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ItemNotificationMessageType"]["ProgressIndicator"])
            hostItemNotificationMessageType = 1;
        else if(notificationType === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ItemNotificationMessageType"]["InformationalMessage"])
            hostItemNotificationMessageType = 0;
        else if(notificationType === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ItemNotificationMessageType"]["ErrorMessage"])
            hostItemNotificationMessageType = 2;
        else if(notificationType === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ItemNotificationMessageType"]["InsightMessage"])
            hostItemNotificationMessageType = 3;
        else
            throw Error.argument("type");
        dataToHost["type"] = hostItemNotificationMessageType
    };
    $h.NotificationMessages._validateKey$p = function(key)
    {
        if(!$h.ScriptHelpers.isNonEmptyString(key))
            throw Error.argument("key");
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(key.length,0,32,"key")
    };
    $h.NotificationMessages._validateDictionary$p = function(dictionary)
    {
        if(!$h.ScriptHelpers.isNonEmptyString(dictionary["type"]))
            throw Error.argument("type");
        if(dictionary["type"] === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ItemNotificationMessageType"]["InformationalMessage"])
        {
            if(!$h.ScriptHelpers.isNonEmptyString(dictionary["icon"]))
                throw Error.argument("icon");
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(dictionary["icon"].length,0,32,"icon");
            if($h.ScriptHelpers.isUndefined(dictionary["persistent"]))
                throw Error.argument("persistent");
            if(!Boolean["isInstanceOfType"](dictionary["persistent"]))
                throw Error.argumentType("persistent",Object["getType"](dictionary["persistent"]),Boolean);
            $h.NotificationMessages._verifyActionDefinitionIsNotDefined$p(dictionary)
        }
        else if(dictionary["type"] === window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ItemNotificationMessageType"]["InsightMessage"])
        {
            if(!$h.ScriptHelpers.isNonEmptyString(dictionary["icon"]))
                throw Error.argument("icon");
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(dictionary["icon"].length,0,32,"icon");
            if(!$h.ScriptHelpers.isUndefined(dictionary["persistent"]))
                throw Error.argument("persistent");
            if(dictionary["actions"])
                $h.NotificationMessages._validateActionsDefinitionBlob$p(dictionary["actions"],dictionary)
        }
        else
        {
            if(!$h.ScriptHelpers.isUndefined(dictionary["icon"]))
                throw Error.argument("icon");
            if(!$h.ScriptHelpers.isUndefined(dictionary["persistent"]))
                throw Error.argument("persistent");
            $h.NotificationMessages._verifyActionDefinitionIsNotDefined$p(dictionary)
        }
        if(!$h.ScriptHelpers.isNonEmptyString(dictionary["message"]))
            throw Error.argument("message");
        window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(dictionary["message"].length,0,150,"message")
    };
    $h.NotificationMessages._validateActionsDefinitionBlob$p = function(actionsDefinitionBlob, notificationParametersDictionary)
    {
        var actionsDefinition = $h.NotificationMessages._extractActionDefinitionDictionary$p(actionsDefinitionBlob);
        if(!actionsDefinition)
            return;
        $h.NotificationMessages._validateActionsDefinitionActionType$p(actionsDefinition,notificationParametersDictionary);
        $h.NotificationMessages._validateActionsDefinitionActionText$p(actionsDefinition)
    };
    $h.NotificationMessages._verifyActionDefinitionIsNotDefined$p = function(notificationParametersDictionary)
    {
        if(!$h.ScriptHelpers.isUndefined(notificationParametersDictionary["actions"]))
            throw Error.argument("actions",window["_u"]["ExtensibilityStrings"]["l_ActionsDefinitionWrongNotificationMessageError_Text"]);
    };
    $h.NotificationMessages._extractActionDefinitionDictionary$p = function(actionsDefinitionBlob)
    {
        var actionsDefinition = null;
        if(Array["isInstanceOfType"](actionsDefinitionBlob))
        {
            var dicArray = actionsDefinitionBlob;
            if(dicArray["length"] === 1)
                actionsDefinition = dicArray[0];
            else if(dicArray["length"] > 1)
                throw Error.argument("actions",window["_u"]["ExtensibilityStrings"]["l_ActionsDefinitionMultipleActionsError_Text"]);
        }
        else
            throw Error.argument("actions",String.format(window["_u"]["ExtensibilityStrings"]["l_InvalidParameterValueError_Text"],"actions"));
        return actionsDefinition
    };
    $h.NotificationMessages._validateActionsDefinitionActionType$p = function(actionsDefinition, notificationParametersDictionary)
    {
        if(!actionsDefinition["actionType"])
            throw Error.argument("actionType",String.format(window["_u"]["ExtensibilityStrings"]["l_NullOrEmptyParameterError_Text"],"actionType"));
        if("showTaskPane" !== actionsDefinition["actionType"])
            throw Error.argument("actionType",window["_u"]["ExtensibilityStrings"]["l_InvalidActionType_Text"]);
        else if(!$h.ScriptHelpers.isNonEmptyString(actionsDefinition["commandId"]))
            throw Error.argument("commandId",String.format(window["_u"]["ExtensibilityStrings"]["l_InvalidCommandIdError_Text"],"commandId"));
    };
    $h.NotificationMessages._validateActionsDefinitionActionText$p = function(actionsDefinition)
    {
        if(!$h.ScriptHelpers.isNonEmptyString(actionsDefinition["actionText"]))
            throw Error.argument("actionText",String.format(window["_u"]["ExtensibilityStrings"]["l_NullOrEmptyParameterError_Text"],"actionText"));
        if(actionsDefinition["actionText"].length > 30)
            throw Error.argument(window["_u"]["ExtensibilityStrings"]["l_ParameterValueTooLongError_Text"],String.format(window["_u"]["ExtensibilityStrings"]["l_ParameterValueTooLongError_Text"],"actionText",30));
    };
    $h.NotificationMessages.prototype.addAsync = function(key, dictionary)
    {
        var args = [];
        for(var $$pai_5 = 2; $$pai_5 < arguments["length"]; ++$$pai_5)
            args[$$pai_5 - 2] = arguments[$$pai_5];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(0,"NotificationMessages.addAsync");
        var commonParameters = $h.CommonParameters.parse(args,false);
        $h.NotificationMessages._validateKey$p(key);
        $h.NotificationMessages._validateDictionary$p(dictionary);
        var dataToHost = {};
        dataToHost = $h.ScriptHelpers.deepClone(dictionary);
        dataToHost["key"] = key;
        $h.NotificationMessages._mapToHostItemNotificationMessageType$p(dataToHost);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(33,dataToHost,null,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
    };
    $h.NotificationMessages.prototype.getAllAsync = function()
    {
        var args = [];
        for(var $$pai_2 = 0; $$pai_2 < arguments["length"]; ++$$pai_2)
            args[$$pai_2] = arguments[$$pai_2];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(0,"NotificationMessages.getAllAsync");
        var commonParameters = $h.CommonParameters.parse(args,true);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(34,null,null,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
    };
    $h.NotificationMessages.prototype.replaceAsync = function(key, dictionary)
    {
        var args = [];
        for(var $$pai_5 = 2; $$pai_5 < arguments["length"]; ++$$pai_5)
            args[$$pai_5 - 2] = arguments[$$pai_5];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(0,"NotificationMessages.replaceAsync");
        var commonParameters = $h.CommonParameters.parse(args,false);
        $h.NotificationMessages._validateKey$p(key);
        $h.NotificationMessages._validateDictionary$p(dictionary);
        var dataToHost = {};
        dataToHost = $h.ScriptHelpers.deepClone(dictionary);
        dataToHost["key"] = key;
        $h.NotificationMessages._mapToHostItemNotificationMessageType$p(dataToHost);
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(35,dataToHost,null,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
    };
    $h.NotificationMessages.prototype.removeAsync = function(key)
    {
        var args = [];
        for(var $$pai_4 = 1; $$pai_4 < arguments["length"]; ++$$pai_4)
            args[$$pai_4 - 1] = arguments[$$pai_4];
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._throwOnMethodCallForInsufficientPermission$i$0(0,"NotificationMessages.removeAsync");
        var commonParameters = $h.CommonParameters.parse(args,false);
        $h.NotificationMessages._validateKey$p(key);
        var dataToHost = {key: key};
        window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._standardInvokeHostMethod$i$0(36,dataToHost,null,commonParameters._asyncContext$p$0,commonParameters._callback$p$0)
    };
    window["Microsoft"]["Office"]["WebExtension"]["OutlookBase"] = function Microsoft_Office_WebExtension_OutlookBase(){};
    window["Microsoft"]["Office"]["WebExtension"]["OutlookBase"]["SeriesTimeJsonConverter"] = function(rawInput)
    {
        if(rawInput && Object["isInstanceOfType"](rawInput))
        {
            var rawDictionary = rawInput;
            if(rawDictionary["seriesTimeJson"])
            {
                var seriesTime = new window["Microsoft"]["Office"]["WebExtension"]["SeriesTime"];
                seriesTime.importFromSeriesTimeJsonObject(rawDictionary["seriesTimeJson"]);
                delete rawDictionary["seriesTimeJson"];
                rawDictionary["seriesTime"] = seriesTime
            }
        }
        return rawInput
    };
    window["Microsoft"]["Office"]["WebExtension"]["OutlookBase"]["CreateAttachmentDetails"] = function(data)
    {
        return new $h.AttachmentDetails(data)
    };
    $h.OutlookErrorManager = function(){};
    $h.OutlookErrorManager.getErrorArgs = function(errorCode)
    {
        if(!$h.OutlookErrorManager._isInitialized$p)
            $h.OutlookErrorManager._initialize$p();
        return OSF.DDA.ErrorCodeManager["getErrorArgs"](errorCode)
    };
    $h.OutlookErrorManager._initialize$p = function()
    {
        $h.OutlookErrorManager._addErrorMessage$p(9e3,"AttachmentSizeExceeded",window["_u"]["ExtensibilityStrings"]["l_AttachmentExceededSize_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9001,"NumberOfAttachmentsExceeded",window["_u"]["ExtensibilityStrings"]["l_ExceededMaxNumberOfAttachments_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9002,"InternalFormatError",window["_u"]["ExtensibilityStrings"]["l_InternalFormatError_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9003,"InvalidAttachmentId",window["_u"]["ExtensibilityStrings"]["l_InvalidAttachmentId_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9004,"InvalidAttachmentPath",window["_u"]["ExtensibilityStrings"]["l_InvalidAttachmentPath_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9005,"CannotAddAttachmentBeforeUpgrade",window["_u"]["ExtensibilityStrings"]["l_CannotAddAttachmentBeforeUpgrade_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9006,"AttachmentDeletedBeforeUploadCompletes",window["_u"]["ExtensibilityStrings"]["l_AttachmentDeletedBeforeUploadCompletes_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9007,"AttachmentUploadGeneralFailure",window["_u"]["ExtensibilityStrings"]["l_AttachmentUploadGeneralFailure_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9008,"AttachmentToDeleteDoesNotExist",window["_u"]["ExtensibilityStrings"]["l_DeleteAttachmentDoesNotExist_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9009,"AttachmentDeleteGeneralFailure",window["_u"]["ExtensibilityStrings"]["l_AttachmentDeleteGeneralFailure_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9010,"InvalidEndTime",window["_u"]["ExtensibilityStrings"]["l_InvalidEndTime_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9011,"HtmlSanitizationFailure",window["_u"]["ExtensibilityStrings"]["l_HtmlSanitizationFailure_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9012,"NumberOfRecipientsExceeded",String.format(window["_u"]["ExtensibilityStrings"]["l_NumberOfRecipientsExceeded_Text"],500));
        $h.OutlookErrorManager._addErrorMessage$p(9013,"NoValidRecipientsProvided",window["_u"]["ExtensibilityStrings"]["l_NoValidRecipientsProvided_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9014,"CursorPositionChanged",window["_u"]["ExtensibilityStrings"]["l_CursorPositionChanged_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9016,"InvalidSelection",window["_u"]["ExtensibilityStrings"]["l_InvalidSelection_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9017,"AccessRestricted","");
        $h.OutlookErrorManager._addErrorMessage$p(9018,"GenericTokenError","");
        $h.OutlookErrorManager._addErrorMessage$p(9019,"GenericSettingsError","");
        $h.OutlookErrorManager._addErrorMessage$p(9020,"GenericResponseError","");
        $h.OutlookErrorManager._addErrorMessage$p(9021,"SaveError",window["_u"]["ExtensibilityStrings"]["l_SaveError_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9022,"MessageInDifferentStoreError",window["_u"]["ExtensibilityStrings"]["l_MessageInDifferentStoreError_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9023,"DuplicateNotificationKey",window["_u"]["ExtensibilityStrings"]["l_DuplicateNotificationKey_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9024,"NotificationKeyNotFound",window["_u"]["ExtensibilityStrings"]["l_NotificationKeyNotFound_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9025,"NumberOfNotificationsExceeded",window["_u"]["ExtensibilityStrings"]["l_NumberOfNotificationsExceeded_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9026,"PersistedNotificationArrayReadError",window["_u"]["ExtensibilityStrings"]["l_PersistedNotificationArrayReadError_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9027,"PersistedNotificationArraySaveError",window["_u"]["ExtensibilityStrings"]["l_PersistedNotificationArraySaveError_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9028,"CannotPersistPropertyInUnsavedDraftError",window["_u"]["ExtensibilityStrings"]["l_CannotPersistPropertyInUnsavedDraftError_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9029,"CanOnlyGetTokenForSavedItem",window["_u"]["ExtensibilityStrings"]["l_CallSaveAsyncBeforeToken_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9030,"APICallFailedDueToItemChange",window["_u"]["ExtensibilityStrings"]["l_APICallFailedDueToItemChange_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9031,"InvalidParameterValueError",window["_u"]["ExtensibilityStrings"]["l_InvalidParameterValueError_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9033,"SetRecurrenceOnInstanceError",window["_u"]["ExtensibilityStrings"]["l_Recurrence_Error_Instance_SetAsync_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9034,"InvalidRecurrenceError",window["_u"]["ExtensibilityStrings"]["l_Recurrence_Error_Properties_Invalid_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9035,"RecurrenceZeroOccurrences",window["_u"]["ExtensibilityStrings"]["l_RecurrenceErrorZeroOccurrences_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9036,"RecurrenceMaxOccurrences",window["_u"]["ExtensibilityStrings"]["l_RecurrenceErrorMaxOccurrences_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9037,"RecurrenceInvalidTimeZone",window["_u"]["ExtensibilityStrings"]["l_RecurrenceInvalidTimeZone_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9038,"InsufficientItemPermissionsError",window["_u"]["ExtensibilityStrings"]["l_Insufficient_Item_Permissions_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9039,"RecurrenceUnsupportedAlternateCalendar",window["_u"]["ExtensibilityStrings"]["l_RecurrenceUnsupportedAlternateCalendar_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9040,"HTTPRequestFailure",window["_u"]["ExtensibilityStrings"]["l_Olk_Http_Error_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9041,"NetworkError",window["_u"]["ExtensibilityStrings"]["l_Internet_Not_Connected_Error_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9042,"InternalServerError",window["_u"]["ExtensibilityStrings"]["l_Internal_Server_Error_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9043,"AttachmentTypeNotSupported",window["_u"]["ExtensibilityStrings"]["l_AttachmentNotSupported_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9044,"InvalidCategory",window["_u"]["ExtensibilityStrings"]["l_Invalid_Category_Error_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9045,"DuplicateCategory",window["_u"]["ExtensibilityStrings"]["l_Duplicate_Category_Error_Text"]);
        $h.OutlookErrorManager._addErrorMessage$p(9046,"ItemNotSaved",window["_u"]["ExtensibilityStrings"]["l_Item_Not_Saved_Error_Text"]);
        $h.OutlookErrorManager._isInitialized$p = true
    };
    $h.OutlookErrorManager._addErrorMessage$p = function(errorCode, errorName, errorMessage)
    {
        OSF.DDA.ErrorCodeManager["addErrorMessage"](errorCode,{
            name: errorName,
            message: errorMessage
        })
    };
    $h.OutlookErrorManager.OutlookErrorCodes = function(){};
    $h.OutlookErrorManager.OsfDdaErrorCodes = function(){};
    $h.PhoneNumber = function(data)
    {
        this.$$d__getPhoneType$p$0 = Function.createDelegate(this,this._getPhoneType$p$0);
        this.$$d__getOriginalPhoneString$p$0 = Function.createDelegate(this,this._getOriginalPhoneString$p$0);
        this.$$d__getPhoneString$p$0 = Function.createDelegate(this,this._getPhoneString$p$0);
        this._data$p$0 = data;
        $h.InitialData._defineReadOnlyProperty$i(this,"phoneString",this.$$d__getPhoneString$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"originalPhoneString",this.$$d__getOriginalPhoneString$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"type",this.$$d__getPhoneType$p$0)
    };
    $h.PhoneNumber.prototype = {
        _data$p$0: null,
        _getPhoneString$p$0: function()
        {
            return this._data$p$0["PhoneString"]
        },
        _getOriginalPhoneString$p$0: function()
        {
            return this._data$p$0["OriginalPhoneString"]
        },
        _getPhoneType$p$0: function()
        {
            return this._data$p$0["Type"]
        }
    };
    $h.TaskSuggestion = function(data)
    {
        this.$$d__getAssignees$p$0 = Function.createDelegate(this,this._getAssignees$p$0);
        this.$$d__getTaskString$p$0 = Function.createDelegate(this,this._getTaskString$p$0);
        this._data$p$0 = data;
        $h.InitialData._defineReadOnlyProperty$i(this,"taskString",this.$$d__getTaskString$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"assignees",this.$$d__getAssignees$p$0)
    };
    $h.TaskSuggestion.prototype = {
        _data$p$0: null,
        _assignees$p$0: null,
        _getTaskString$p$0: function()
        {
            return this._data$p$0["TaskString"]
        },
        _getAssignees$p$0: function()
        {
            if(!this._assignees$p$0)
            {
                var $$t_1 = this;
                this._assignees$p$0 = $h.Entities._getExtractedObjects$i($h.EmailAddressDetails,this._data$p$0,"Assignees",function(data)
                {
                    return $h.EmailAddressDetails._createFromEmailUserDictionary$i(data)
                })
            }
            return this._assignees$p$0
        }
    };
    $h.UserProfile = function(data)
    {
        this.$$d__getUserProfileType$p$0 = Function.createDelegate(this,this._getUserProfileType$p$0);
        this.$$d__getTimeZone$p$0 = Function.createDelegate(this,this._getTimeZone$p$0);
        this.$$d__getEmailAddress$p$0 = Function.createDelegate(this,this._getEmailAddress$p$0);
        this.$$d__getDisplayName$p$0 = Function.createDelegate(this,this._getDisplayName$p$0);
        this._data$p$0 = data;
        $h.InitialData._defineReadOnlyProperty$i(this,"displayName",this.$$d__getDisplayName$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"emailAddress",this.$$d__getEmailAddress$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"timeZone",this.$$d__getTimeZone$p$0);
        $h.InitialData._defineReadOnlyProperty$i(this,"accountType",this.$$d__getUserProfileType$p$0)
    };
    $h.UserProfile.prototype = {
        _data$p$0: null,
        _getUserProfileType$p$0: function()
        {
            return this._data$p$0.get__userProfileType$i$0()
        },
        _getDisplayName$p$0: function()
        {
            return this._data$p$0.get__userDisplayName$i$0()
        },
        _getEmailAddress$p$0: function()
        {
            return this._data$p$0.get__userEmailAddress$i$0()
        },
        _getTimeZone$p$0: function()
        {
            return this._data$p$0.get__userTimeZone$i$0()
        }
    };
    $h.OutlookDispid = function(){};
    $h.OutlookDispid.prototype = {
        owaOnlyMethod: 0,
        getInitialData: 1,
        getUserIdentityToken: 2,
        loadCustomProperties: 3,
        saveCustomProperties: 4,
        ewsRequest: 5,
        displayNewAppointmentForm: 7,
        displayMessageForm: 8,
        displayAppointmentForm: 9,
        displayReplyForm: 10,
        displayReplyAllForm: 11,
        getCallbackToken: 12,
        bodySetSelectedDataAsync: 13,
        getBodyTypeAsync: 14,
        getRecipientsAsync: 15,
        addFileAttachmentAsync: 16,
        setSubjectAsync: 17,
        getSubjectAsync: 18,
        addItemAttachmentAsync: 19,
        removeAttachmentAsync: 20,
        setRecipientsAsync: 21,
        addRecipientsAsync: 22,
        bodyPrependAsync: 23,
        getTimeAsync: 24,
        setTimeAsync: 25,
        getLocationAsync: 26,
        setLocationAsync: 27,
        getSelectedDataAsync: 28,
        setSelectedDataAsync: 29,
        displayReplyFormWithAttachments: 30,
        displayReplyAllFormWithAttachments: 31,
        saveAsync: 32,
        addNotficationMessageAsync: 33,
        getAllNotficationMessagesAsync: 34,
        replaceNotficationMessageAsync: 35,
        removeNotficationMessageAsync: 36,
        getBodyAsync: 37,
        setBodyAsync: 38,
        appCommands1: 39,
        registerConsentAsync: 40,
        close: 41,
        closeApp: 42,
        displayPersonaCardAsync: 43,
        displayNewMessageForm: 44,
        navigateToModuleAsync: 45,
        eventCompleted: 94,
        closeContainer: 97,
        getInitializationContextAsync: 99,
        appendOnSendAsync: 100,
        moveToFolder: 101,
        getRecurrenceAsync: 103,
        setRecurrenceAsync: 104,
        getFromAsync: 107,
        getSharedPropertiesAsync: 108,
        messageParent: 144,
        addBase64FileAttachmentAsync: 148,
        getAttachmentsAsync: 149,
        getAttachmentContentAsync: 150,
        getInternetHeadersAsync: 151,
        setInternetHeadersAsync: 152,
        removeInternetHeadersAsync: 153,
        getEnhancedLocationsAsync: 154,
        addEnhancedLocationsAsync: 155,
        removeEnhancedLocationsAsync: 156,
        getCategoriesAsync: 157,
        addCategoriesAsync: 158,
        removeCategoriesAsync: 159,
        getMasterCategoriesAsync: 160,
        addMasterCategoriesAsync: 161,
        removeMasterCategoriesAsync: 162,
        logTelemetry: 163,
        getItemIdAsync: 164,
        trackCtq: 400,
        recordTrace: 401,
        recordDataPoint: 402,
        windowOpenOverrideHandler: 403,
        saveSettingsRequest: 404
    };
    $h.OutlookDispid["registerEnum"]("$h.OutlookDispid",false);
    $h.RequestState = function(){};
    $h.RequestState.prototype = {
        unsent: 0,
        opened: 1,
        headersReceived: 2,
        loading: 3,
        done: 4
    };
    $h.RequestState["registerEnum"]("$h.RequestState",false);
    $h.CommonParameters = function(options, callback, asyncContext)
    {
        this._options$p$0 = options;
        this._callback$p$0 = callback;
        this._asyncContext$p$0 = asyncContext
    };
    $h.CommonParameters.parse = function(args, isCallbackRequired, tryLegacy)
    {
        var legacyParameters;
        var $$t_8,
            $$t_9;
        if(tryLegacy && ($$t_9 = $h.CommonParameters._tryParseLegacy$p(args,$$t_8 = {val: legacyParameters}),legacyParameters = $$t_8["val"],$$t_9))
            return legacyParameters;
        var argsLength = args["length"];
        var options = null;
        var callback = null;
        var asyncContext = null;
        if(argsLength === 1)
            if($h.CommonParameters._argIsFunction$p(args[0]))
                callback = args[0];
            else if(Object["isInstanceOfType"](args[0]))
                options = args[0];
            else
                throw Error.argumentType();
        else if(argsLength === 2)
        {
            if(!Object["isInstanceOfType"](args[0]))
                throw Error.argument("options");
            if(!$h.CommonParameters._argIsFunction$p(args[1]))
                throw Error.argument("callback");
            options = args[0];
            callback = args[1]
        }
        else if(argsLength)
            throw Error.parameterCount(window["_u"]["ExtensibilityStrings"]["l_ParametersNotAsExpected_Text"]);
        if(isCallbackRequired && !callback)
            throw Error.argumentNull("callback");
        if(options && !$h.ScriptHelpers.isNullOrUndefined(options["asyncContext"]))
            asyncContext = options["asyncContext"];
        return new $h.CommonParameters(options,callback,asyncContext)
    };
    $h.CommonParameters._tryParseLegacy$p = function(args, commonParameters)
    {
        commonParameters["val"] = null;
        var argsLength = args["length"];
        var callback = null;
        var userContext = null;
        if(!argsLength || argsLength > 2)
            return false;
        if(!$h.CommonParameters._argIsFunction$p(args[0]))
            return false;
        callback = args[0];
        if(argsLength > 1)
            userContext = args[1];
        commonParameters["val"] = new $h.CommonParameters(null,callback,userContext);
        return true
    };
    $h.CommonParameters._argIsFunction$p = function(arg)
    {
        return typeof arg === "function"
    };
    $h.CommonParameters.prototype = {
        _options$p$0: null,
        _callback$p$0: null,
        _asyncContext$p$0: null,
        get_options: function()
        {
            return this._options$p$0
        },
        get_callback: function()
        {
            return this._callback$p$0
        },
        get_asyncContext: function()
        {
            return this._asyncContext$p$0
        }
    };
    $h.ShouldRunNewCodeForFlags = function(){};
    $h.EwsRequest = function(userContext)
    {
        $h.EwsRequest["initializeBase"](this,[userContext])
    };
    $h.EwsRequest.prototype = {
        readyState: 1,
        status: 0,
        statusText: null,
        onreadystatechange: null,
        responseText: null,
        get__statusCode$i$1: function()
        {
            return this.status
        },
        set__statusCode$i$1: function(value)
        {
            this.status = value;
            return value
        },
        get__statusDescription$i$1: function()
        {
            return this.statusText
        },
        set__statusDescription$i$1: function(value)
        {
            this.statusText = value;
            return value
        },
        get__requestState$i$1: function()
        {
            return this.readyState
        },
        set__requestState$i$1: function(value)
        {
            this.readyState = value;
            return value
        },
        get_hasOnReadyStateChangeCallback: function()
        {
            return!$h.ScriptHelpers.isNullOrUndefined(this.onreadystatechange)
        },
        get__response$i$1: function()
        {
            return this.responseText
        },
        set__response$i$1: function(value)
        {
            this.responseText = value;
            return value
        },
        send: function(data)
        {
            this._checkSendConditions$i$1();
            if($h.ScriptHelpers.isNullOrUndefined(data))
                this._throwInvalidStateException$i$1();
            this._sendRequest$i$0(5,"EwsRequest",{body: data})
        },
        _callOnReadyStateChangeCallback$i$1: function()
        {
            if(!$h.ScriptHelpers.isNullOrUndefined(this.onreadystatechange))
                this.onreadystatechange()
        },
        _parseExtraResponseData$i$1: function(response){},
        executeExtraFailedResponseSteps: function(){}
    };
    $h.InitialData = function(data)
    {
        this._data$p$0 = data;
        this._permissionLevel$p$0 = this._calculatePermissionLevel$p$0()
    };
    $h.InitialData._defineReadOnlyProperty$i = function(o, methodName, getter)
    {
        var propertyDescriptor = {
                get: getter,
                configurable: false
            };
        window["Object"]["defineProperty"](o,methodName,propertyDescriptor)
    };
    $h.InitialData.prototype = {
        _toRecipients$p$0: null,
        _ccRecipients$p$0: null,
        _attachments$p$0: null,
        _resources$p$0: null,
        _entities$p$0: null,
        _selectedEntities$p$0: null,
        _data$p$0: null,
        _permissionLevel$p$0: 0,
        get__isRestIdSupported$i$0: function()
        {
            return this._data$p$0["isRestIdSupported"]
        },
        get__itemId$i$0: function()
        {
            return this._data$p$0["id"]
        },
        get__itemClass$i$0: function()
        {
            return this._data$p$0["itemClass"]
        },
        get__dateTimeCreated$i$0: function()
        {
            return new Date(this._data$p$0["dateTimeCreated"])
        },
        get__dateTimeModified$i$0: function()
        {
            return new Date(this._data$p$0["dateTimeModified"])
        },
        get__dateTimeSent$i$0: function()
        {
            return new Date(this._data$p$0["dateTimeSent"])
        },
        get__subject$i$0: function()
        {
            this._throwOnRestrictedPermissionLevel$i$0();
            return this._data$p$0["subject"]
        },
        get__normalizedSubject$i$0: function()
        {
            this._throwOnRestrictedPermissionLevel$i$0();
            return this._data$p$0["normalizedSubject"]
        },
        get__internetMessageId$i$0: function()
        {
            return this._data$p$0["internetMessageId"]
        },
        get__conversationId$i$0: function()
        {
            return this._data$p$0["conversationId"]
        },
        get__sender$i$0: function()
        {
            this._throwOnRestrictedPermissionLevel$i$0();
            var sender = this._data$p$0["sender"];
            return $h.ScriptHelpers.isNullOrUndefined(sender) ? null : new $h.EmailAddressDetails(sender)
        },
        get__from$i$0: function()
        {
            this._throwOnRestrictedPermissionLevel$i$0();
            var from = this._data$p$0["from"];
            return $h.ScriptHelpers.isNullOrUndefined(from) ? null : new $h.EmailAddressDetails(from)
        },
        get__to$i$0: function()
        {
            this._throwOnRestrictedPermissionLevel$i$0();
            if(null === this._toRecipients$p$0)
                this._toRecipients$p$0 = this._createEmailAddressDetails$p$0("to");
            return this._toRecipients$p$0
        },
        get__cc$i$0: function()
        {
            this._throwOnRestrictedPermissionLevel$i$0();
            if(null === this._ccRecipients$p$0)
                this._ccRecipients$p$0 = this._createEmailAddressDetails$p$0("cc");
            return this._ccRecipients$p$0
        },
        get__attachments$i$0: function()
        {
            this._throwOnRestrictedPermissionLevel$i$0();
            if(null === this._attachments$p$0)
                this._attachments$p$0 = $h.AttachmentDetails.createFromJsonArray(this._data$p$0["attachments"]);
            return this._attachments$p$0
        },
        get__ewsUrl$i$0: function()
        {
            return this._data$p$0["ewsUrl"]
        },
        get__restUrl$i$0: function()
        {
            return this._data$p$0["restUrl"]
        },
        get__marketplaceAssetId$i$0: function()
        {
            return this._data$p$0["marketplaceAssetId"]
        },
        get__extensionId$i$0: function()
        {
            return this._data$p$0["extensionId"]
        },
        get__marketplaceContentMarket$i$0: function()
        {
            return this._data$p$0["marketplaceContentMarket"]
        },
        get__consentMetadata$i$0: function()
        {
            return this._data$p$0["consentMetadata"]
        },
        get__isRead$i$0: function()
        {
            return this._data$p$0["isRead"]
        },
        get__isFromSharedFolder$i$0: function()
        {
            return!!this._data$p$0["isFromSharedFolder"] && this._data$p$0["isFromSharedFolder"]
        },
        get__shouldRunNewCodeForFlags$i$0: function()
        {
            if(this._data$p$0["shouldRunNewCodeForFlags"])
                return this._data$p$0["shouldRunNewCodeForFlags"];
            return 0
        },
        get__endNodeUrl$i$0: function()
        {
            return this._data$p$0["endNodeUrl"]
        },
        get__entryPointUrl$i$0: function()
        {
            return this._data$p$0["entryPointUrl"]
        },
        get__start$i$0: function()
        {
            return new Date(this._data$p$0["start"])
        },
        get__end$i$0: function()
        {
            return new Date(this._data$p$0["end"])
        },
        get__location$i$0: function()
        {
            return this._data$p$0["location"]
        },
        get__userProfileType$i$0: function()
        {
            return this._data$p$0["userProfileType"]
        },
        get__resources$i$0: function()
        {
            this._throwOnRestrictedPermissionLevel$i$0();
            if(null === this._resources$p$0)
                this._resources$p$0 = this._createEmailAddressDetails$p$0("resources");
            return this._resources$p$0
        },
        get__organizer$i$0: function()
        {
            this._throwOnRestrictedPermissionLevel$i$0();
            var organizer = this._data$p$0["organizer"];
            return $h.ScriptHelpers.isNullOrUndefined(organizer) ? null : new $h.EmailAddressDetails(organizer)
        },
        get__recurrence$i$0: function()
        {
            this._throwOnRestrictedPermissionLevel$i$0();
            return this._data$p$0["recurrence"]
        },
        get__seriesId$i$0: function()
        {
            this._throwOnRestrictedPermissionLevel$i$0();
            return this._data$p$0["seriesId"]
        },
        get__userDisplayName$i$0: function()
        {
            return this._data$p$0["userDisplayName"]
        },
        get__userEmailAddress$i$0: function()
        {
            return this._data$p$0["userEmailAddress"]
        },
        get__userTimeZone$i$0: function()
        {
            return this._data$p$0["userTimeZone"]
        },
        get__timeZoneOffsets$i$0: function()
        {
            return this._data$p$0["timeZoneOffsets"]
        },
        get__hostVersion$i$0: function()
        {
            return this._data$p$0["hostVersion"]
        },
        get__owaView$i$0: function()
        {
            return this._data$p$0["owaView"]
        },
        get__overrideWindowOpen$i$0: function()
        {
            return this._data$p$0["overrideWindowOpen"]
        },
        _getEntities$i$0: function()
        {
            if(!this._entities$p$0)
                this._entities$p$0 = new $h.Entities(this._data$p$0["entities"],this._data$p$0["filteredEntities"],this.get__dateTimeSent$i$0(),this._permissionLevel$p$0);
            return this._entities$p$0
        },
        _getSelectedEntities$i$0: function()
        {
            if(!this._selectedEntities$p$0)
                this._selectedEntities$p$0 = new $h.Entities(this._data$p$0["selectedEntities"],null,this.get__dateTimeSent$i$0(),this._permissionLevel$p$0);
            return this._selectedEntities$p$0
        },
        _getEntitiesByType$i$0: function(entityType)
        {
            var entites = this._getEntities$i$0();
            return entites._getByType$i$0(entityType)
        },
        _getFilteredEntitiesByName$i$0: function(name)
        {
            var entities = this._getEntities$i$0();
            return entities._getFilteredEntitiesByName$i$0(name)
        },
        _getRegExMatches$i$0: function()
        {
            if(!this._data$p$0["regExMatches"])
                return null;
            return this._data$p$0["regExMatches"]
        },
        _getSelectedRegExMatches$i$0: function()
        {
            if(!this._data$p$0["selectedRegExMatches"])
                return null;
            return this._data$p$0["selectedRegExMatches"]
        },
        _getRegExMatchesByName$i$0: function(regexName)
        {
            var regexMatches = this._getRegExMatches$i$0();
            if(!regexMatches || !regexMatches[regexName])
                return null;
            return regexMatches[regexName]
        },
        _throwOnRestrictedPermissionLevel$i$0: function()
        {
            window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i(this._permissionLevel$p$0)
        },
        _createEmailAddressDetails$p$0: function(key)
        {
            var to = this._data$p$0[key];
            if($h.ScriptHelpers.isNullOrUndefined(to))
                return[];
            var recipients = [];
            for(var i = 0; i < to["length"]; i++)
                if(!$h.ScriptHelpers.isNullOrUndefined(to[i]))
                    recipients[i] = new $h.EmailAddressDetails(to[i]);
            return recipients
        },
        _calculatePermissionLevel$p$0: function()
        {
            var HostReadItem = 1;
            var HostReadWriteMailbox = 2;
            var HostReadWriteItem = 3;
            var permissionLevelFromHost = this._data$p$0["permissionLevel"];
            if($h.ScriptHelpers.isUndefined(this._permissionLevel$p$0))
                return 0;
            switch(permissionLevelFromHost)
            {
                case HostReadItem:
                    return 1;
                case HostReadWriteItem:
                    return 2;
                case HostReadWriteMailbox:
                    return 3;
                default:
                    return 0
            }
        }
    };
    $h._loadDictionaryRequest = function(createResultObject, dictionaryName, callback, userContext)
    {
        $h._loadDictionaryRequest["initializeBase"](this,[userContext]);
        this._createResultObject$p$1 = createResultObject;
        this._dictionaryName$p$1 = dictionaryName;
        this._callback$p$1 = callback
    };
    $h._loadDictionaryRequest.prototype = {
        _dictionaryName$p$1: null,
        _createResultObject$p$1: null,
        _callback$p$1: null,
        handleResponse: function(response)
        {
            if(response["wasSuccessful"])
            {
                var value = response[this._dictionaryName$p$1];
                var responseData = window["JSON"]["parse"](value);
                this.createAsyncResult(this._createResultObject$p$1(responseData),0,0,null)
            }
            else
                this.createAsyncResult(null,1,9020,response["errorMessage"]);
            this._callback$p$1(this._asyncResult$p$0)
        }
    };
    $h.ProxyRequestBase = function(userContext)
    {
        $h.ProxyRequestBase["initializeBase"](this,[userContext])
    };
    $h.ProxyRequestBase.prototype = {
        handleResponse: function(response)
        {
            if(!response["wasProxySuccessful"])
            {
                this.set__statusCode$i$1(500);
                this.set__statusDescription$i$1("Error");
                var errorMessage = response["errorMessage"];
                this.set__response$i$1(errorMessage);
                this.createAsyncResult(null,1,9020,errorMessage)
            }
            else
            {
                this.set__statusCode$i$1(response["statusCode"]);
                this.set__statusDescription$i$1(response["statusDescription"]);
                this.set__response$i$1(response["body"]);
                this.createAsyncResult(this.get__response$i$1(),0,0,null)
            }
            this._parseExtraResponseData$i$1(response);
            this._cycleReadyStateFromHeadersReceivedToLoadingToDone$i$1()
        },
        _throwInvalidStateException$i$1: function()
        {
            throw Error.create("DOMException",{
                code: 11,
                message: "INVALID_STATE_ERR"
            });
        },
        _cycleReadyStateFromHeadersReceivedToLoadingToDone$i$1: function()
        {
            var $$t_0 = this;
            this._changeReadyState$i$1(2,function()
            {
                $$t_0._changeReadyState$i$1(3,function()
                {
                    $$t_0._changeReadyState$i$1(4,null)
                })
            })
        },
        _changeReadyState$i$1: function(state, nextStep)
        {
            this.set__requestState$i$1(state);
            var $$t_2 = this;
            window.setTimeout(function()
            {
                try
                {
                    $$t_2._callOnReadyStateChangeCallback$i$1()
                }
                finally
                {
                    if(!$h.ScriptHelpers.isNullOrUndefined(nextStep))
                        nextStep()
                }
            },0)
        },
        _checkSendConditions$i$1: function()
        {
            if(this.get__requestState$i$1() !== 1)
                this._throwInvalidStateException$i$1();
            if(this._isSent$p$0)
                this._throwInvalidStateException$i$1()
        }
    };
    $h.RequestBase = function(userContext)
    {
        this._userContext$p$0 = userContext
    };
    $h.RequestBase.prototype = {
        _isSent$p$0: false,
        _asyncResult$p$0: null,
        _userContext$p$0: null,
        get_asyncResult: function()
        {
            return this._asyncResult$p$0
        },
        _sendRequest$i$0: function(dispid, methodName, dataToSend)
        {
            this._isSent$p$0 = true;
            var $$t_5 = this;
            window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.invokeHostMethod(dispid,dataToSend,function(resultCode, response)
            {
                if(resultCode)
                    $$t_5.createAsyncResult(null,1,9017,String.format(window["_u"]["ExtensibilityStrings"]["l_InternalProtocolError_Text"],resultCode));
                else
                    $$t_5.handleResponse(response)
            })
        },
        createAsyncResult: function(value, errorCode, detailedErrorCode, errorDescription)
        {
            this._asyncResult$p$0 = window["OSF"]["DDA"]["OutlookAppOm"]._instance$p.createAsyncResult(value,errorCode,detailedErrorCode,this._userContext$p$0,errorDescription)
        }
    };
    $h.SaveDictionaryRequest = function(callback, userContext)
    {
        $h.SaveDictionaryRequest["initializeBase"](this,[userContext]);
        if(!$h.ScriptHelpers.isNullOrUndefined(callback))
            this._callback$p$1 = callback
    };
    $h.SaveDictionaryRequest.prototype = {
        _callback$p$1: null,
        handleResponse: function(response)
        {
            if(response["wasSuccessful"])
                this.createAsyncResult(null,0,0,null);
            else
                this.createAsyncResult(null,1,9020,response["errorMessage"]);
            if(!$h.ScriptHelpers.isNullOrUndefined(this._callback$p$1))
                this._callback$p$1(this._asyncResult$p$0)
        }
    };
    $h.ScriptHelpers = function(){};
    $h.ScriptHelpers.isNull = function(value)
    {
        return null === value
    };
    $h.ScriptHelpers.isNullOrUndefined = function(value)
    {
        return $h.ScriptHelpers.isNull(value) || $h.ScriptHelpers.isUndefined(value)
    };
    $h.ScriptHelpers.isUndefined = function(value)
    {
        return value === undefined
    };
    $h.ScriptHelpers.dictionaryContainsKey = function(obj, keyName)
    {
        return Object["isInstanceOfType"](obj) ? keyName in obj : false
    };
    $h.ScriptHelpers.isNonEmptyString = function(value)
    {
        if(!value)
            return false;
        return String["isInstanceOfType"](value)
    };
    $h.ScriptHelpers.deepClone = function(obj)
    {
        return window["JSON"]["parse"](window["JSON"]["stringify"](obj))
    };
    $h.ScriptHelpers.isValueTrue = function(value)
    {
        if(!$h.ScriptHelpers.isNullOrUndefined(value))
            return value["toString"]().toLowerCase() === "true";
        return false
    };
    $h.ScriptHelpers.validateCategoriesArray = function(categories)
    {
        if($h.ScriptHelpers.isNullOrUndefined(categories))
            throw Error.argument("categories");
        if(!Array["isInstanceOfType"](categories))
            throw Error.argumentType("categories",Object["getType"](categories),Array);
        if(!categories["length"])
            throw Error.argument("categories");
        for(var i = 0; i < categories["length"]; i++)
            if(!$h.ScriptHelpers.isNonEmptyString(categories[i]) || categories[i].length > 255)
                throw Error.argument("categories");
    };
    window["OSF"]["DDA"]["OutlookAppOm"]["registerClass"]("OSF.DDA.OutlookAppOm");
    window["OSF"]["DDA"]["Settings"]["registerClass"]("OSF.DDA.Settings");
    $h.AdditionalGlobalParameters["registerClass"]("$h.AdditionalGlobalParameters");
    $h.ItemBase["registerClass"]("$h.ItemBase");
    $h.Item["registerClass"]("$h.Item",$h.ItemBase);
    $h.Appointment["registerClass"]("$h.Appointment",$h.Item);
    $h.ComposeItem["registerClass"]("$h.ComposeItem",$h.ItemBase);
    $h.AppointmentCompose["registerClass"]("$h.AppointmentCompose",$h.ComposeItem);
    $h.AttachmentDetails["registerClass"]("$h.AttachmentDetails");
    $h.Body["registerClass"]("$h.Body");
    $h.Categories["registerClass"]("$h.Categories");
    $h.ComposeFrom["registerClass"]("$h.ComposeFrom");
    $h.InternetHeaders["registerClass"]("$h.InternetHeaders");
    $h.ComposeBody["registerClass"]("$h.ComposeBody",$h.Body);
    $h.ComposeRecipient["registerClass"]("$h.ComposeRecipient");
    $h.ComposeRecurrence["registerClass"]("$h.ComposeRecurrence");
    $h.ComposeLocation["registerClass"]("$h.ComposeLocation");
    $h.ComposeSubject["registerClass"]("$h.ComposeSubject");
    $h.ComposeTime["registerClass"]("$h.ComposeTime");
    $h.Contact["registerClass"]("$h.Contact");
    $h.CustomProperties["registerClass"]("$h.CustomProperties");
    $h.Diagnostics["registerClass"]("$h.Diagnostics");
    $h.EmailAddressDetails["registerClass"]("$h.EmailAddressDetails");
    $h.EnhancedLocation["registerClass"]("$h.EnhancedLocation");
    $h.Entities["registerClass"]("$h.Entities");
    window["Microsoft"]["Office"]["WebExtension"]["SeriesTime"]["registerClass"]("Microsoft.Office.WebExtension.SeriesTime");
    $h.MasterCategories["registerClass"]("$h.MasterCategories");
    $h.Message["registerClass"]("$h.Message",$h.Item);
    $h.MeetingRequest["registerClass"]("$h.MeetingRequest",$h.Message);
    $h.MeetingSuggestion["registerClass"]("$h.MeetingSuggestion");
    $h._extractedDate["registerClass"]("$h._extractedDate");
    $h._preciseDate["registerClass"]("$h._preciseDate",$h._extractedDate);
    $h._relativeDate["registerClass"]("$h._relativeDate",$h._extractedDate);
    $h.MessageCompose["registerClass"]("$h.MessageCompose",$h.ComposeItem);
    $h.NotificationMessages["registerClass"]("$h.NotificationMessages");
    window["Microsoft"]["Office"]["WebExtension"]["OutlookBase"]["registerClass"]("Microsoft.Office.WebExtension.OutlookBase");
    $h.PhoneNumber["registerClass"]("$h.PhoneNumber");
    $h.TaskSuggestion["registerClass"]("$h.TaskSuggestion");
    $h.UserProfile["registerClass"]("$h.UserProfile");
    $h.CommonParameters["registerClass"]("$h.CommonParameters");
    $h.RequestBase["registerClass"]("$h.RequestBase");
    $h.ProxyRequestBase["registerClass"]("$h.ProxyRequestBase",$h.RequestBase);
    $h.EwsRequest["registerClass"]("$h.EwsRequest",$h.ProxyRequestBase);
    $h.InitialData["registerClass"]("$h.InitialData");
    $h._loadDictionaryRequest["registerClass"]("$h._loadDictionaryRequest",$h.RequestBase);
    $h.SaveDictionaryRequest["registerClass"]("$h.SaveDictionaryRequest",$h.RequestBase);
    window["OSF"]["DDA"]["OutlookAppOm"].asyncMethodTimeoutKeyName = "__timeout__";
    window["OSF"]["DDA"]["OutlookAppOm"].ewsIdOrEmailParamName = "ewsIdOrEmail";
    window["OSF"]["DDA"]["OutlookAppOm"].moduleParamName = "module";
    window["OSF"]["DDA"]["OutlookAppOm"].queryStringParamName = "queryString";
    window["OSF"]["DDA"]["OutlookAppOm"]._maxRecipients$p = 100;
    window["OSF"]["DDA"]["OutlookAppOm"]._maxSubjectLength$p = 255;
    window["OSF"]["DDA"]["OutlookAppOm"].maxBodyLength = 32768;
    window["OSF"]["DDA"]["OutlookAppOm"]._maxLocationLength$p = 255;
    window["OSF"]["DDA"]["OutlookAppOm"]._maxEwsRequestSize$p = 1e6;
    window["OSF"]["DDA"]["OutlookAppOm"].executeMethodName = "ExecuteMethod";
    window["OSF"]["DDA"]["OutlookAppOm"].getInitialDataMethodName = "GetInitialData";
    window["OSF"]["DDA"]["OutlookAppOm"].standardInvokeHostMethodErrorCodeKey = "errorCode";
    window["OSF"]["DDA"]["OutlookAppOm"].standardInvokeHostMethodErrorKey = "error";
    window["OSF"]["DDA"]["OutlookAppOm"].standardInvokeHostMethodDiagnosticsKey = "diagnostics";
    window["OSF"]["DDA"]["OutlookAppOm"].outlookAsyncResponseWasSuccessfulKey = "wasSuccessful";
    window["OSF"]["DDA"]["OutlookAppOm"].outlookAsyncResponseErrorMessageKey = "errorMessage";
    window["OSF"]["DDA"]["OutlookAppOm"].itemIdParameterName = "itemId";
    window["OSF"]["DDA"]["OutlookAppOm"].restVersionParameterName = "restVersion";
    window["OSF"]["DDA"]["OutlookAppOm"]._instance$p = null;
    $h.AdditionalGlobalParameters.itemNumberKey = "itemNumber";
    $h.AdditionalGlobalParameters.actionsDefinitionKey = "actions";
    $h.AttachmentConstants.maxAttachmentNameLength = 255;
    $h.AttachmentConstants.maxUrlLength = 2048;
    $h.AttachmentConstants.maxItemIdLength = 200;
    $h.AttachmentConstants.maxRemoveIdLength = 200;
    $h.AttachmentConstants.attachmentParameterName = "attachments";
    $h.AttachmentConstants.attachmentTypeParameterName = "type";
    $h.AttachmentConstants.attachmentUrlParameterName = "url";
    $h.AttachmentConstants.base64StringParameterName = "base64String";
    $h.AttachmentConstants.attachmentItemIdParameterName = "itemId";
    $h.AttachmentConstants.attachmentNameParameterName = "name";
    $h.AttachmentConstants.attachmentIsInlineParameterName = "isInline";
    $h.AttachmentConstants.attachmentTypeFileName = "file";
    $h.AttachmentConstants.attachmentTypeItemName = "item";
    $h.AttachmentConstants.attachmentIdParameterName = "id";
    $h.AttachmentDetails._attachmentTypeMap$p = [window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["AttachmentType"]["File"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["AttachmentType"]["Item"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["AttachmentType"]["Cloud"]];
    $h.Body.coercionTypeParameterName = "coercionType";
    $h.InternetHeaders.internetHeadersLimit = 998;
    $h.ComposeRecipient.displayNameLengthLimit = 255;
    $h.ComposeRecipient.recipientsLimit = 100;
    $h.ComposeRecipient.totalRecipientsLimit = 500;
    $h.ComposeRecipient.addressParameterName = "address";
    $h.ComposeRecipient.nameParameterName = "name";
    $h.ComposeRecurrence.startDateKey = "startDate";
    $h.ComposeRecurrence.endDateKey = "endDate";
    $h.ComposeRecurrence.startTimeKey = "startTime";
    $h.ComposeRecurrence.endTimeKey = "endTime";
    $h.ComposeRecurrence.recurrenceTypeKey = "recurrenceType";
    $h.ComposeRecurrence.seriesTimeKey = "seriesTime";
    $h.ComposeRecurrence.seriesTimeJsonKey = "seriesTimeJson";
    $h.ComposeRecurrence.recurrenceTimeZoneKey = "recurrenceTimeZone";
    $h.ComposeRecurrence.recurrenceTimeZoneName = "name";
    $h.ComposeRecurrence.recurrencePropertiesKey = "recurrenceProperties";
    $h.ComposeRecurrence.intervalKey = "interval";
    $h.ComposeRecurrence.daysKey = "days";
    $h.ComposeRecurrence.dayOfMonthKey = "dayOfMonth";
    $h.ComposeRecurrence.dayOfWeekKey = "dayOfWeek";
    $h.ComposeRecurrence.weekNumberKey = "weekNumber";
    $h.ComposeRecurrence.monthKey = "month";
    $h.ComposeRecurrence.firstDayOfWeekKey = "firstDayOfWeek";
    $h.ComposeLocation.locationKey = "location";
    $h.ComposeLocation.maximumLocationLength = 255;
    $h.ComposeSubject.maximumSubjectLength = 255;
    $h.ComposeTime.timeTypeName = "TimeProperty";
    $h.ComposeTime.timeDataName = "time";
    $h.Diagnostics.outlookAppName = "Outlook";
    $h.Diagnostics.outlookWebAppName = "OutlookWebApp";
    $h.Diagnostics.outlookIOSAppName = "OutlookIOS";
    $h.Diagnostics.outlookAndroidAppName = "OutlookAndroid";
    $h.EmailAddressDetails._emptyString$p = "";
    $h.EmailAddressDetails._responseTypeMap$p = [window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ResponseType"]["None"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ResponseType"]["Organizer"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ResponseType"]["Tentative"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ResponseType"]["Accepted"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["ResponseType"]["Declined"]];
    $h.EmailAddressDetails._recipientTypeMap$p = [window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecipientType"]["Other"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecipientType"]["DistributionList"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecipientType"]["User"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RecipientType"]["ExternalUser"]];
    $h.Entities._allEntityKeys$p = ["Addresses","EmailAddresses","Urls","PhoneNumbers","TaskSuggestions","MeetingSuggestions","Contacts","FlightReservations","ParcelDeliveries"];
    window["Microsoft"]["Office"]["WebExtension"]["SeriesTime"].startYearKey = "startYear";
    window["Microsoft"]["Office"]["WebExtension"]["SeriesTime"].startMonthKey = "startMonth";
    window["Microsoft"]["Office"]["WebExtension"]["SeriesTime"].startDayKey = "startDay";
    window["Microsoft"]["Office"]["WebExtension"]["SeriesTime"].endYearKey = "endYear";
    window["Microsoft"]["Office"]["WebExtension"]["SeriesTime"].endMonthKey = "endMonth";
    window["Microsoft"]["Office"]["WebExtension"]["SeriesTime"].endDayKey = "endDay";
    window["Microsoft"]["Office"]["WebExtension"]["SeriesTime"].noEndDateKey = "noEndDate";
    window["Microsoft"]["Office"]["WebExtension"]["SeriesTime"].startTimeMinKey = "startTimeMin";
    window["Microsoft"]["Office"]["WebExtension"]["SeriesTime"].durationMinKey = "durationMin";
    $h.ReplyConstants.htmlBodyKeyName = "htmlBody";
    $h.EmailAddressConstants.maxSmtpLength = 571;
    $h.CategoriesConstants.categoriesCharacterLimit = 255;
    $h.AsyncConstants.optionsKeyName = "options";
    $h.AsyncConstants.callbackKeyName = "callback";
    $h.AsyncConstants.asyncResultKeyName = "asyncResult";
    $h.ApiTelemetryCode.success = 0;
    $h.ApiTelemetryCode.noResponseDictionary = -900;
    $h.ApiTelemetryCode.noErrorCodeForStandardInvokeMethod = -901;
    $h.ApiTelemetryCode.genericProxyError = -902;
    $h.ApiTelemetryCode.genericLegacyApiError = -903;
    $h.ApiTelemetryCode.genericUnknownError = -904;
    $h.Item.destFolderParameterName = "destinationFolder";
    $h.MasterCategories._colorPresets$i = [window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["None"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset0"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset1"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset2"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset3"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset4"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset5"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset6"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset7"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset8"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset9"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset10"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset11"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset12"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset13"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset14"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset15"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset16"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset17"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset18"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset19"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset20"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset21"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset22"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset23"],window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["CategoryColor"]["Preset24"]];
    $h.MeetingSuggestionTimeDecoder._baseDate$p = new Date("0001-01-01T00:00:00Z");
    $h.NotificationMessages.maximumKeyLength = 32;
    $h.NotificationMessages.maximumIconLength = 32;
    $h.NotificationMessages.maximumMessageLength = 150;
    $h.NotificationMessages.maximumActionTextLength = 30;
    $h.NotificationMessages.notificationsKeyParameterName = "key";
    $h.NotificationMessages.notificationsTypeParameterName = "type";
    $h.NotificationMessages.notificationsIconParameterName = "icon";
    $h.NotificationMessages.notificationsMessageParameterName = "message";
    $h.NotificationMessages.notificationsPersistentParameterName = "persistent";
    $h.NotificationMessages.notificationsActionsDefinitionParameterName = "actions";
    $h.NotificationMessages.notificationsActionTypeParameterName = "actionType";
    $h.NotificationMessages.notificationsActionTextParameterName = "actionText";
    $h.NotificationMessages.notificationsActionCommandIdParameterName = "commandId";
    $h.NotificationMessages.notificationsActionContextDataParameterName = "contextData";
    $h.NotificationMessages.notificationsActionShowTaskPaneActionId = "showTaskPane";
    $h.OutlookErrorManager.errorNameKey = "name";
    $h.OutlookErrorManager.errorMessageKey = "message";
    $h.OutlookErrorManager._isInitialized$p = false;
    $h.OutlookErrorManager.OutlookErrorCodes.ooeInvalidDataFormat = 2006;
    $h.OutlookErrorManager.OutlookErrorCodes.attachmentSizeExceeded = 9e3;
    $h.OutlookErrorManager.OutlookErrorCodes.numberOfAttachmentsExceeded = 9001;
    $h.OutlookErrorManager.OutlookErrorCodes.internalFormatError = 9002;
    $h.OutlookErrorManager.OutlookErrorCodes.invalidAttachmentId = 9003;
    $h.OutlookErrorManager.OutlookErrorCodes.invalidAttachmentPath = 9004;
    $h.OutlookErrorManager.OutlookErrorCodes.cannotAddAttachmentBeforeUpgrade = 9005;
    $h.OutlookErrorManager.OutlookErrorCodes.attachmentDeletedBeforeUploadCompletes = 9006;
    $h.OutlookErrorManager.OutlookErrorCodes.attachmentUploadGeneralFailure = 9007;
    $h.OutlookErrorManager.OutlookErrorCodes.attachmentToDeleteDoesNotExist = 9008;
    $h.OutlookErrorManager.OutlookErrorCodes.attachmentDeleteGeneralFailure = 9009;
    $h.OutlookErrorManager.OutlookErrorCodes.invalidEndTime = 9010;
    $h.OutlookErrorManager.OutlookErrorCodes.htmlSanitizationFailure = 9011;
    $h.OutlookErrorManager.OutlookErrorCodes.numberOfRecipientsExceeded = 9012;
    $h.OutlookErrorManager.OutlookErrorCodes.noValidRecipientsProvided = 9013;
    $h.OutlookErrorManager.OutlookErrorCodes.cursorPositionChanged = 9014;
    $h.OutlookErrorManager.OutlookErrorCodes.invalidSelection = 9016;
    $h.OutlookErrorManager.OutlookErrorCodes.accessRestricted = 9017;
    $h.OutlookErrorManager.OutlookErrorCodes.genericTokenError = 9018;
    $h.OutlookErrorManager.OutlookErrorCodes.genericSettingsError = 9019;
    $h.OutlookErrorManager.OutlookErrorCodes.genericResponseError = 9020;
    $h.OutlookErrorManager.OutlookErrorCodes.saveError = 9021;
    $h.OutlookErrorManager.OutlookErrorCodes.messageInDifferentStoreError = 9022;
    $h.OutlookErrorManager.OutlookErrorCodes.duplicateNotificationKey = 9023;
    $h.OutlookErrorManager.OutlookErrorCodes.notificationKeyNotFound = 9024;
    $h.OutlookErrorManager.OutlookErrorCodes.numberOfNotificationsExceeded = 9025;
    $h.OutlookErrorManager.OutlookErrorCodes.persistedNotificationArrayReadError = 9026;
    $h.OutlookErrorManager.OutlookErrorCodes.persistedNotificationArraySaveError = 9027;
    $h.OutlookErrorManager.OutlookErrorCodes.cannotPersistPropertyInUnsavedDraftError = 9028;
    $h.OutlookErrorManager.OutlookErrorCodes.callSaveAsyncBeforeToken = 9029;
    $h.OutlookErrorManager.OutlookErrorCodes.apiCallFailedDueToItemChange = 9030;
    $h.OutlookErrorManager.OutlookErrorCodes.invalidParameterValueError = 9031;
    $h.OutlookErrorManager.OutlookErrorCodes.setRecurrenceOnInstance = 9033;
    $h.OutlookErrorManager.OutlookErrorCodes.invalidRecurrence = 9034;
    $h.OutlookErrorManager.OutlookErrorCodes.recurrenceZeroOccurrences = 9035;
    $h.OutlookErrorManager.OutlookErrorCodes.recurrenceMaxOccurrences = 9036;
    $h.OutlookErrorManager.OutlookErrorCodes.recurrenceInvalidTimeZone = 9037;
    $h.OutlookErrorManager.OutlookErrorCodes.insufficientItemPermissions = 9038;
    $h.OutlookErrorManager.OutlookErrorCodes.recurrenceUnsupportedAlternateCalendar = 9039;
    $h.OutlookErrorManager.OutlookErrorCodes.olkHTTPError = 9040;
    $h.OutlookErrorManager.OutlookErrorCodes.olkInternetNotConnectedError = 9041;
    $h.OutlookErrorManager.OutlookErrorCodes.olkInternalServerError = 9042;
    $h.OutlookErrorManager.OutlookErrorCodes.attachmentTypeNotSupported = 9043;
    $h.OutlookErrorManager.OutlookErrorCodes.olkInvalidCategoryError = 9044;
    $h.OutlookErrorManager.OutlookErrorCodes.duplicateCategoryError = 9045;
    $h.OutlookErrorManager.OutlookErrorCodes.itemNotSavedError = 9046;
    $h.OutlookErrorManager.OsfDdaErrorCodes.ooeCoercionTypeNotSupported = 1e3;
    $h.OutlookErrorManager.OsfDdaErrorCodes.ooeOperationNotSupported = 5e3;
    $h.CommonParameters.asyncContextKeyName = "asyncContext";
    $h.ShouldRunNewCodeForFlags.saveCustomProperties = 1;
    $h.InitialData.userProfileTypeKey = "userProfileType";
    $h.ScriptHelpers.emptyString = "";
    OSF.DDA.ErrorCodeManager.initializeErrorMessages(Strings.OfficeOM);
    if(appContext.get_appName() == OSF.AppName.OutlookWebApp || appContext.get_appName() == OSF.AppName.OutlookIOS || appContext.get_appName() == OSF.AppName.OutlookAndroid)
        this._settings = this._initializeSettings(appContext,false);
    else
        this._settings = this._initializeSettings(false);
    appContext.appOM = new OSF.DDA.OutlookAppOm(appContext,this._webAppState.wnd,appReady);
    if(appContext.get_appName() == OSF.AppName.Outlook || appContext.get_appName() == OSF.AppName.OutlookWebApp || appContext.get_appName() == OSF.AppName.OutlookIOS || appContext.get_appName() == OSF.AppName.OutlookAndroid)
        OSF.DDA.DispIdHost.addEventSupport(appContext.appOM,new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.ItemChanged,Microsoft.Office.WebExtension.EventType.OfficeThemeChanged]))
}
