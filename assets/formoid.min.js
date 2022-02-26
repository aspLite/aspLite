jQuery(function ($) {

    var Formoid = (function () {

        var API_URL = ('https:' == location.protocol ? 'https:' : 'http:') + '//formoid.net/api/push';

        var $ajax = (function () {
            var ie = (/MSIE (\d+)\./.exec(navigator.userAgent) || [0, 0])[1];
            if (8 == ie || (9 == ie && 'file:' != location.protocol)) {
                return function (url, settings) {
                    var xdr = new XDomainRequest(), defer = $.Deferred();
                    xdr.open(settings.type, url);
                    xdr.onload = function () {
                        defer.resolve(this.responseText);
                    };
                    xdr.onerror = function () {
                        defer.reject();
                    };
                    xdr.send(settings.data);
                    return defer;
                };
            } else {
                $.support.cors = true;
                return $.ajax;
            }
        })();

        var prop = function (name, args) {
            name = '__' + name + '__';
            if (args.length) {
                this[name] = args[0];
                return this;
            }
            return this[name];
        };

        var attachMethods = function (obj, context, methods) {
            $.each(methods, function (i, method) {
                obj[method] = function () {
                    return context[method].apply(context, arguments);
                };
            });
            return obj;
        };

        var Form = function (settings) {
            settings = settings || {};
            this.__email__ = settings.email || '';
            this.__title__ = settings.title || '';
            this.__data__ = settings.data || [];
            this.__recaptcha__ = settings.recaptcha || '';
        };

        Form.prototype.email = function (value) {
            return prop.call(this, 'email', arguments);
        };

        Form.prototype.title = function (value) {
            return prop.call(this, 'title', arguments);
        };

        Form.prototype.recaptcha_token = function (value) {
            return prop.call(this, 'recaptcha_token', arguments);
        };

        Form.prototype.data = function (value) {
            return prop.call(this, 'data', arguments);
        };

        Form.prototype.getCaptchaToken = function() {
            var recaptcha = this.__recaptcha__;
            var def = $.Deferred();
            if ( recaptcha ) {
                grecaptcha.ready(function() {
                    try {
                        grecaptcha.execute(recaptcha, {action: 'homepage'})
                        .then(function(token) {
                            def.resolve(token);
                        });
                    } catch(error) {
                        console.error(error);
                        def.resolve(null);
                    }
                  });
            } else { def.resolve(null); }
            return def;
        }

        Form.prototype.send = function (data, beforeStart) {
            var _this = this;
            var defer = attachMethods($.Deferred(), this, ['email', 'title', 'data', 'send']);
            var recaptchaToken = this.__recaptcha__ ? this.getCaptchaToken() : $.when();
            if (beforeStart) {
                beforeStart.call(this, defer);
                if ('pending' != defer.state())
                    return defer;
            }

            recaptchaToken
            .then(function(token) {
                if (token) {
                    _this.__rcph_token__ = token;
                }
            })
            .always(function() {
                var objData = {
                   email: _this.__email__,
                   form: {
                       title: _this.__title__,
                       data: (arguments.length ? data : _this.__data__)
                   }
                }

                if ( _this.__rcph_token__ ) objData.recaptcha_token = _this.__rcph_token__;

                $ajax(API_URL, {
                   type: 'POST',
                   data: JSON.stringify(objData)
               }).done(function (responseText) {
                   try {
                       var data = JSON.parse(responseText);
                       if (data.error) defer.reject(data.error);
                       else defer.resolve(data.response);
                   } catch (e) {
                       defer.reject('Incorrect server response.');
                   }
               }).fail(function () {
                   var error = 'Failed to query the server. ';
                   if ('onLine' in navigator && !navigator.onLine)
                       error += 'No connection to the Internet.';
                   else error += 'Check the connection and try again.';
                   defer.reject(error);
               }); 
            });
            return defer;
        };

        return {
            Form: function (settings) {
                return new Form(settings);
            }
        }

    })();

    var isValidEmail = function (input) {
        var isValid = true,
            value = $(input).val();

        if (value) {
            isValid = /^([^@]+?)@(([a-z0-9]-*)*[a-z0-9]+\.)+([a-z0-9]+)$/i.test(value);
        }

        return isValid;
    };

    $('[data-form-type="formoid"]').each(function () {
        var form,
            $this = $(this),
            $form = $this.is('form') ? $this : $this.find('form'),
            action = $form.attr('action'),
            $alert = $this.find('[data-form-alert]'),
            $title = $this.is('[data-form-title]') ? $this : $this.find('[data-form-title]'),
            $submit = $this.find('[type="submit"]'),
            $inputArr = $form.find('[data-form-field]'),
            alertSuccess = $form.find('[data-form-alert]').length ? $alert.html() : $alert.find('[data-form-alert-success]').html(),
            $captcha = $this.find('[data-form-captcha]');

        // on change form inputs if input[type=file].files[0]size>1mb disable sumbit button and show form alert
        $form.change(function (event) {
            if (event.target.type === 'file') {
                if (event.target.files[0].size > 1000000) {
                    $alert.removeAttr('hidden');
                    $alert.removeClass('alert-success').addClass('alert-danger');
                    $alert.html('File size must be less than 1mb');
                    $submit.addClass('btn-loading').prop('disabled', true);
                }
            }
        });

        $inputArr.each(function (index, elem) {
            $(elem).on("focus", function () {
                $alert.attr('hidden', 'hidden');
                $alert.addClass('alert-success').removeClass('alert-danger');
                $alert.html('');
                $submit.removeClass('btn-loading').prop('disabled', false);
            });
        });

        // event on form submit
        $form.submit(function (event) {
            if (action !== 'https://mobirise.com/') return;
            event.preventDefault();
            if ($submit.hasClass('btn-loading')) return;

            var data = [],
                inputs = $inputArr;

            $form.addClass('form-active');
            $submit.addClass('btn-loading').prop('disabled', true);
            $alert.html('');
            form = form || Formoid.Form({
                email: $this.find('[data-form-email]').val(),
                title: $title.attr('data-form-title') || $title.text(),
                recaptcha: $this.find('[data-form-captcha]').val() || false
            });

            function parseInputs(inputs) {
                var def = $.Deferred(),
                    arrDeferred = [];

                inputs.each(function (index, input) {
                    var $input = $(input),
                        typeInput = $input.attr('type'),
                        objDeferred = $.Deferred();

                    arrDeferred.push(objDeferred);

                    // valide email
                    if ($input.attr('name') === "email") {
                        if (!isValidEmail(input)) {
                            $submit.removeClass('btn-loading').prop('disabled', true);
                            def.reject(new Error('Form is not valid'));
                            return false;
                        }
                    }

                    switch (typeInput) {
                        case 'file': {
                            var reader = new FileReader(),
                                name = $input.attr('data-form-field') || $input.attr('name'),
                                files = $input[0].files[0];

                            reader.onloadend = function () {
                                data.push([
                                    name,
                                    reader.result
                                ]);
                                objDeferred.resolve();
                            };
                            reader.onerror = function () {
                                $alert.html(reader.error);
                                def.reject(reader.error);
                            };

                            if (files) {
                                reader.readAsDataURL(files);
                            }
                        }
                            break;
                        case 'checkbox': {
                            data.push([
                                $input.attr('data-form-field') || $input.attr('name'),
                                $input.prop("checked") ? $input.val() : 'No'
                            ]);
                            objDeferred.resolve();
                        }
                            break;
                        case 'radio': {
                            if ($input.prop("checked")) {
                                data.push([
                                    $input.attr('data-form-field') || $input.attr('name'),
                                    $input.val()
                                ]);
                            }
                            objDeferred.resolve();
                        }
                            break;
                        default: {
                            data.push([
                                $input.attr('data-form-field') || $input.attr('name'),
                                $input.val()
                            ]);
                        }
                        objDeferred.resolve();
                    }
                });

                return def.resolve(arrDeferred);
            }

            parseInputs(inputs).then(function (arrDef) {
                $.when.apply(this, arrDef)
                .then(function() {
                    form.send(data)
                        .always(function () {
                            $submit.removeClass('btn-loading').prop('disabled', false);
                            $alert.removeAttr('hidden');
                        })
                        .then(function (message) {
                            inputs.each(function () {
                                if ($(this).prop('checked') && $(this).attr('type') === 'checkbox') $(this).removeAttr('checked');
                                else $(this).val('');
                            });
                            $form.removeClass('form-active');
                            $alert.text(alertSuccess || message);
                            $alert.removeClass('alert-danger').addClass('alert-success');
                        })
                        .fail(function (error) {
                            $alert.text(error);
                            $alert.removeClass('alert-success').addClass('alert-danger');
                        });
                }, function (err) {
                    $alert.html(err.message);
                    $alert.removeClass('alert-success').addClass('alert-danger');
                    $alert.removeAttr('hidden');
                });
            });
        });
    });
});