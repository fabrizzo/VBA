Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit

Option Compare Database
Private Sub Comeback_Click()
DoCmd.Close acForm, Me.Name, acSaveNo
DoCmd.OpenForm "Forma_st", acNormal
End Sub
Private Sub Form_Activate()
Dim ctl As Control
For Each ctl In Me.Controls
With ctl
     If ctl.ControlType = acComboBox Then
     Select Case Left(ctl.Name, 6)
     Case "Form1U"
            .SetFocus
            .Text = "Form1"
     Case "Spisok"
            If ctl.Name = "Spisok_Field1" Then
            .SetFocus
            .Text = "Form1.f1_3num"
            Else
            If ctl.Name = "Spisok_Field2" Then
            .SetFocus
            .Text = "Form1.f1_4num"
            Else
            .SetFocus
            .RowSourceType = "Value List"
            .AddItem "", 0
            .AddItem "Form1.f1_1kod", 1
            .AddItem "Form1.f1_3num", 2
            .AddItem "Form1.f1_4num", 3
            .AddItem "Form1.ovd", 4
            .AddItem "Form1.f1_3v", 5
            .AddItem "Form1.f1_13s", 6
            .AddItem "Form1.f1_13z", 7
            .AddItem "Form1.f1_13ch", 8
            .AddItem "Form1.f1_13p1_1", 9
            .AddItem "Form1.f1_13p1_2", 10
            .AddItem "Form1.f1_13p1_3", 11
            .AddItem "Form1.f1_13p1_4", 12
            .AddItem "Form1.f1_13p1_5", 13
            .AddItem "Form1.f1_14s", 14
            .AddItem "Form1.f1_14z", 15
            .AddItem "Form1.f1_14ch", 16
            .AddItem "Form1.f1_14p1_1", 17
            .AddItem "Form1.f1_14p1_2", 18
            .AddItem "Form1.f1_16s", 19
            .AddItem "Form1.f1_n", 20
            .AddItem "Form1.f1_d", 21
            .AddItem "Form1.f1_1v", 22
            .AddItem "Form1.f1_2", 23
            .AddItem "Form1.f1_5num", 24
            .AddItem "Form1.f1_5d", 25
            .AddItem "Form1.f1_7kod", 26
            .AddItem "Form1.f1_7d", 27
            .AddItem "Form1.f1_8", 28
            .AddItem "Form1.f1_9", 29
            .AddItem "Form1.f1_91", 30
            .AddItem "Form1.f1_901_1", 31
            .AddItem "Form1.f1_901_2", 32
            .AddItem "Form1.f1_10", 33
            .AddItem "Form1.f1_101", 34
            .AddItem "Form1.f1_102", 35
            .AddItem "Form1.f1_103", 36
            .AddItem "Form1.f1_103_1", 37
            .AddItem "Form1.f1_104_2", 38
            .AddItem "Form1.f1_104_1", 39
            .AddItem "Form1.f1_104", 40
            .AddItem "Form1.f1_105_1", 41
            .AddItem "Form1.f1_105", 42
            .AddItem "Form1.f1_11d", 43
            .AddItem "Form1.f1_11", 44
            .AddItem "Form1.f1_111n", 45
            .AddItem "Form1.f1_111", 46
            .AddItem "Form1.f1_12k", 47
            .AddItem "Form1.f1_12d", 48
            .AddItem "Form1.f1_12v", 49
            .AddItem "Form1.f1_13sn", 50
            .AddItem "Form1.f1_13zn", 51
            .AddItem "Form1.f1_13chn", 52
            .AddItem "Form1.f1_13p1n_1", 53
            .AddItem "Form1.f1_13p1n_2", 54
            .AddItem "Form1.f1_13p1n_3", 55
            .AddItem "Form1.f1_13p1n_4", 56
            .AddItem "Form1.f1_13p1n_5", 57
            .AddItem "Form1.f1_14sn", 58
            .AddItem "Form1.f1_14zn", 59
            .AddItem "Form1.f1_14chn", 60
            .AddItem "Form1.f1_14p1n_1", 61
            .AddItem "Form1.f1_14p1n_2", 62
            .AddItem "Form1.f1_15", 63
            .AddItem "Form1.f1_16", 64
            .AddItem "Form1.f1_17", 65
            .AddItem "Form1.f1_181", 66
            .AddItem "Form1.f1_18", 67
            .AddItem "Form1.f1_190", 68
            .AddItem "Form1.f1_19", 69
            .AddItem "Form1.f1_191", 70
            .AddItem "Form1.f1_192", 71
            .AddItem "Form1.f1_192u", 72
            .AddItem "Form1.f1_192m", 73
            .AddItem "Form1.f1_192d", 74
            .AddItem "Form1.f1_192k", 75
            .AddItem "Form1.f1_192kv", 76
            .AddItem "Form1.f1_193", 77
            .AddItem "Form1.f1_21", 78
            .AddItem "Form1.f1_2111", 79
            .AddItem "Form1.f1_211", 80
            .AddItem "Form1.f1_212", 81
            .AddItem "Form1.f1_213", 82
            .AddItem "Form1.f1_20", 83
            .AddItem "Form1.f1_22", 84
            .AddItem "Form1.f1_221", 85
            .AddItem "Form1.f1_23_1", 86
            .AddItem "Form1.f1_23_2", 87
            .AddItem "Form1.f1_23_3", 88
            .AddItem "Form1.f1_231_1", 89
            .AddItem "Form1.f1_231_2", 90
            .AddItem "Form1.f1_231_3", 91
            .AddItem "Form1.f1_24", 92
            .AddItem "Form1.f1_241", 93
            .AddItem "Form1.f1_242", 94
            .AddItem "Form1.f1_243", 95
            .AddItem "Form1.f1_244", 96
            .AddItem "Form1.f1_245", 97
            .AddItem "Form1.f1_253", 98
            .AddItem "Form1.f1_2540", 99
            .AddItem "Form1.f1_254", 100
            .AddItem "Form1.f1_255k", 101
            .AddItem "Form1.f1_255g", 102
            .AddItem "Form1.f1_255m", 103
            .AddItem "Form1.f1_2560", 104
            .AddItem "Form1.f1_256", 105
            .AddItem "Form1.f1_257k", 106
            .AddItem "Form1.f1_257g", 107
            .AddItem "Form1.f1_257m", 108
            .AddItem "Form1.f1_2580", 109
            .AddItem "Form1.f1_258", 110
            .AddItem "Form1.f1_259k", 111
            .AddItem "Form1.f1_259g", 112
            .AddItem "Form1.f1_259m", 113
            .AddItem "Form1.f1_2600", 114
            .AddItem "Form1.f1_260", 115
            .AddItem "Form1.f1_261k", 116
            .AddItem "Form1.f1_261g", 117
            .AddItem "Form1.f1_261m", 118
            .AddItem "Form1.f1_26_1", 119
            .AddItem "Form1.f1_26_2", 120
            .AddItem "Form1.f1_26_3", 121
            .AddItem "Form1.f1_26_4", 122
            .AddItem "Form1.f1_27_1", 123
            .AddItem "Form1.f1_27_2", 124
            .AddItem "Form1.f1_27_3", 125
            .AddItem "Form1.f1_27_4", 126
            .AddItem "Form1.f1_28", 127
            .AddItem "Form1.f1_281", 128
            .AddItem "Form1.f1_282", 129
            .AddItem "Form1.f1_283", 130
            .AddItem "Form1.f1_29", 131
            .AddItem "Form1.f1_291_1", 132
            .AddItem "Form1.f1_291_2", 133
            .AddItem "Form1.f1_291_3", 134
            .AddItem "Form1.f1_291_4", 135
            .AddItem "Form1.f1_30v", 136
            .AddItem "Form1.f1_30p", 137
            .AddItem "Form1.f1_30r", 138
            .AddItem "Form1.f1_301v", 139
            .AddItem "Form1.f1_301p", 140
            .AddItem "Form1.f1_301r", 141
            .AddItem "Form1.f1_304v", 142
            .AddItem "Form1.f1_304p", 143
            .AddItem "Form1.f1_304r", 144
            .AddItem "Form1.f1_307v", 145
            .AddItem "Form1.f1_307p", 146
            .AddItem "Form1.f1_307r", 147
            .AddItem "Form1.f1_31", 148
            .AddItem "Form1.f1_32", 149
            .AddItem "Form1.f1_321", 150
            .AddItem "Form1.f1_33_1", 151
            .AddItem "Form1.f1_33_2", 152
            .AddItem "Form1.f1_33_3", 153
            .AddItem "Form1.f1_33_4", 154
            .AddItem "Form1.f1_34", 155
            .AddItem "Form1.f1_341", 156
            .AddItem "Form1.f1_342", 157
            .AddItem "Form1.f1_343", 158
            .AddItem "Form1.f1_35", 159
            .AddItem "Form1.f1_35_1", 160
            .AddItem "Form1.f1_35_2", 161
            .AddItem "Form1.f1_35_3", 162
            .AddItem "Form1.f1_35_4", 163
            .AddItem "Form1.f1_352", 164
            .AddItem "Form1.f1_352_1", 165
            .AddItem "Form1.f1_352_2", 166
            .AddItem "Form1.f1_352_3", 167
            .AddItem "Form1.f1_352_4", 168
            .AddItem "Form1.f1_36", 169
            .AddItem "Form1.f1_361", 170
            .AddItem "Form1.f1_36_1", 171
            .AddItem "Form1.f1_37", 172
            .AddItem "Form1.f1_37_1", 173
            .AddItem "Form1.f1_38_1", 174
            .AddItem "Form1.f1_38_2", 175
            .AddItem "Form1.f1_38_3", 176
            .AddItem "Form1.f1_40", 177
            .AddItem "Form1.f1_41", 178
            .AddItem "Form1.f1_2620", 179
            .AddItem "Form1.f1_262", 180
            .AddItem "Form1.f1_262k", 181
            .AddItem "Form1.f1_262g", 182
            .AddItem "Form1.f1_262m", 183
            .AddItem "Form1.f1_264", 184
            .AddItem "Form1.f1_265", 185
            .AddItem "Form1.f1_401", 186
            .AddItem "Form11.f11_25k", 187
            .AddItem "Form11.f11_25d", 188
            .AddItem "Form11.f11_n", 189
            .AddItem "Form11.f11_d", 190
            .AddItem "Form11.f11_1v", 191
            .AddItem "Form11.f11_2", 192
            .AddItem "Form11.f11_6kod", 193
            .AddItem "Form11.f11_6d", 194
            .AddItem "Form11.f11_7s", 195
            .AddItem "Form11.f11_7z", 196
            .AddItem "Form11.f11_7ch", 197
            .AddItem "Form11.f11_7p1_1", 198
            .AddItem "Form11.f11_7p1_2", 199
            .AddItem "Form11.f11_7p1_3", 200
            .AddItem "Form11.f11_7p1_4", 201
            .AddItem "Form11.f11_7p1_5", 202
            .AddItem "Form11.f11_701s", 203
            .AddItem "Form11.f11_701z", 204
            .AddItem "Form11.f11_701c", 205
            .AddItem "Form11.f11_701p_1", 206
            .AddItem "Form11.f11_701p_2", 207
            .AddItem "Form11.f11_703", 208
            .AddItem "Form11.f11_9_1", 209
            .AddItem "Form11.f11_9_2", 210
            .AddItem "Form11.f11_911", 211
            .AddItem "Form11.f11_91", 212
            .AddItem "Form11.f11_10", 213
            .AddItem "Form11.f11_101", 214
            .AddItem "Form11.f11_102", 215
            .AddItem "Form11.f11_103", 216
            .AddItem "Form11.f11_104", 217
            .AddItem "Form11.f11_11e", 218
            .AddItem "Form11.f11_11", 219
            .AddItem "Form11.f11_111e", 220
            .AddItem "Form11.f11_111", 221
            .AddItem "Form11.f11_12_1", 222
            .AddItem "Form11.f11_12_2", 223
            .AddItem "Form11.f11_12_3", 224
            .AddItem "Form11.f11_13_1", 225
            .AddItem "Form11.f11_13_2", 226
            .AddItem "Form11.f11_14_1", 227
            .AddItem "Form11.f11_14_2", 228
            .AddItem "Form11.f11_14_3", 229
            .AddItem "Form11.f11_15", 230
            .AddItem "Form11.f11_151", 231
            .AddItem "Form11.f11_152", 232
            .AddItem "Form11.f11_153", 233
            .AddItem "Form11.f11_154", 234
            .AddItem "Form11.f11_155", 235
            .AddItem "Form11.f11_16", 236
            .AddItem "Form11.f11_17_1", 237
            .AddItem "Form11.f11_17_2", 238
            .AddItem "Form11.f11_17_3", 239
            .AddItem "Form11.f11_18_1", 240
            .AddItem "Form11.f11_18_2", 241
            .AddItem "Form11.f11_181", 242
            .AddItem "Form11.f11_182_1", 243
            .AddItem "Form11.f11_182_2", 244
            .AddItem "Form11.f11_19_1", 245
            .AddItem "Form11.f11_19_2", 246
            .AddItem "Form11.f11_20_1", 247
            .AddItem "Form11.f11_20_2", 248
            .AddItem "Form11.f11_201_1", 249
            .AddItem "Form11.f11_201_2", 250
            .AddItem "Form11.f11_21_1", 251
            .AddItem "Form11.f11_21_2", 252
            .AddItem "Form11.f11_22_1", 253
            .AddItem "Form11.f11_22_2", 254
            .AddItem "Form11.f11_23", 255
            .AddItem "Form11.f11_231", 256
            .AddItem "Form11.f11_232", 257
            .AddItem "Form11.f11_233", 258
            .AddItem "Form11.f11_248", 259
            .AddItem "Form11.f11_24", 260
            .AddItem "Form11.f11_241", 261
            .AddItem "Form11.f11_242", 262
            .AddItem "Form11.f11_243", 263
            .AddItem "Form11.f11_249", 264
            .AddItem "Form11.f11_244", 265
            .AddItem "Form11.f11_245", 266
            .AddItem "Form11.f11_246", 267
            .AddItem "Form11.f11_247", 268
            .AddItem "Form11.f11_258k", 269
            .AddItem "Form11.f11_258d", 270
            .AddItem "Form11.f11_25f", 271
            .AddItem "Form11.f11_25i", 272
            .AddItem "Form11.f11_25o", 273
            .AddItem "Form11.f11_252", 274
            .AddItem "Form11.f11_251", 275
            .AddItem "Form11.f11_25n", 276
            .AddItem "Form11.f11_253", 277
            .AddItem "Form11.f11_254", 278
            .AddItem "Form11.f11_255", 279
            .AddItem "Form11.f11_256", 280
            .AddItem "Form11.f11_257", 281
            .AddItem "Form11.f11_26_1", 282
            .AddItem "Form11.f11_26_2", 283
            .AddItem "Form11.f11_27_1", 284
            .AddItem "Form11.f11_27_2", 285
            .AddItem "Form11.f11_28", 286
            .AddItem "Form11.f11_281", 287
            .AddItem "Form11.f11_282", 288
            .AddItem "Form11.f11_283", 289
            .AddItem "Form11.f11_284", 290
            .AddItem "Form11.f11_285", 291
            .AddItem "Form11.f11_286", 292
            .AddItem "Form11.f11_287", 293
            .AddItem "Form11.f11_288", 294
            .AddItem "Form11.f11_289", 295
            .AddItem "Form11.f11_29_1", 296
            .AddItem "Form11.f11_29_2", 297
            .AddItem "Form11.f11_29_3", 298
            .AddItem "Form11.f11_30_1", 299
            .AddItem "Form11.f11_30_2", 300
            .AddItem "Form11.f11_31", 301
            .AddItem "Form11.f11_311", 302
            .AddItem "Form11.f11_312", 303
            .AddItem "Form11.f11_313", 304
            .AddItem "Form11.f11_314", 305
            .AddItem "Form11.f11_315", 306
            .AddItem "Form11.f11_32_1", 307
            .AddItem "Form11.f11_32_2", 308
            .AddItem "Form11.f11_32_3", 309
            .AddItem "Form11.f11_3211_1", 310
            .AddItem "Form11.f11_3211_2", 311
            .AddItem "Form11.f11_33", 312
            .AddItem "Form11.f11_332", 313
            .AddItem "Form11.f11_333", 314
            .AddItem "Form11.f11_34", 315
            .AddItem "Form11.f11_341", 316
            .AddItem "Form11.f11_342", 317
            .AddItem "Form11.f11_343", 318
            .AddItem "Form11.f11_350", 319
            .AddItem "Form11.f11_362", 320
            .AddItem "Form11.f11_35", 321
            .AddItem "Form11.f11_351", 322
            .AddItem "Form11.f11_352", 323
            .AddItem "Form11.f11_35n", 324
            .AddItem "Form11.f11_3530", 325
            .AddItem "Form11.f11_363", 326
            .AddItem "Form11.f11_353", 327
            .AddItem "Form11.f11_354", 328
            .AddItem "Form11.f11_355", 329
            .AddItem "Form11.f11_353n", 330
            .AddItem "Form11.f11_3560", 331
            .AddItem "Form11.f11_364", 332
            .AddItem "Form11.f11_356", 333
            .AddItem "Form11.f11_357", 334
            .AddItem "Form11.f11_358", 335
            .AddItem "Form11.f11_356n", 336
            .AddItem "Form11.f11_3590", 337
            .AddItem "Form11.f11_365", 338
            .AddItem "Form11.f11_359", 339
            .AddItem "Form11.f11_360", 340
            .AddItem "Form11.f11_361", 341
            .AddItem "Form11.f11_359n", 342
            .AddItem "Form11.f11_36_1", 343
            .AddItem "Form11.f11_36_2", 344
            .AddItem "Form11.f11_38", 345
            .AddItem "Form11.f11_382", 346
            .AddItem "Form11.f11_383", 347
            .AddItem "Form11.f11_386", 348
            .AddItem "Form11.f11_385", 349
            .AddItem "Form11.f11_388", 350
            .AddItem "Form11.f11_390", 351
            .AddItem "Form11.f11_39", 352
            .AddItem "Form11.f11_39k", 353
            .AddItem "Form11.f11_39g", 354
            .AddItem "Form11.f11_39m", 355
            .AddItem "Form11.f11_3910", 356
            .AddItem "Form11.f11_391", 357
            .AddItem "Form11.f11_391k", 358
            .AddItem "Form11.f11_391g", 359
            .AddItem "Form11.f11_391m", 360
            .AddItem "Form11.f11_3920", 361
            .AddItem "Form11.f11_392", 362
            .AddItem "Form11.f11_392k", 363
            .AddItem "Form11.f11_392g", 364
            .AddItem "Form11.f11_392m", 365
            .AddItem "Form11.f11_3930", 366
            .AddItem "Form11.f11_393", 367
            .AddItem "Form11.f11_393k", 368
            .AddItem "Form11.f11_393g", 369
            .AddItem "Form11.f11_393m", 370
            .AddItem "Form11.f11_3940", 371
            .AddItem "Form11.f11_394", 372
            .AddItem "Form11.f11_394k", 373
            .AddItem "Form11.f11_394g", 374
            .AddItem "Form11.f11_394m", 375
            .AddItem "Form11.f11_331", 376
            .AddItem "Form11.f11_37", 377
            .AddItem "Form12.f12_n", 378
            .AddItem "Form12.f12_d", 379
            .AddItem "Form12.f12_1v", 380
            .AddItem "Form12.f12_6kod", 381
            .AddItem "Form12.f12_6d", 382
            .AddItem "Form12.f12_7s", 383
            .AddItem "Form12.f12_7z", 384
            .AddItem "Form12.f12_7ch", 385
            .AddItem "Form12.f12_7p1_1", 386
            .AddItem "Form12.f12_7p1_2", 387
            .AddItem "Form12.f12_7p1_3", 388
            .AddItem "Form12.f12_7p1_4", 389
            .AddItem "Form12.f12_7p1_5", 390
            .AddItem "Form12.f12_8s", 391
            .AddItem "Form12.f12_8z", 392
            .AddItem "Form12.f12_8ch", 393
            .AddItem "Form12.f12_8p1_1", 394
            .AddItem "Form12.f12_8p1_2", 395
            .AddItem "Form12.f12_121", 396
            .AddItem "Form12.f12_13", 397
            .AddItem "Form12.f12_131f", 398
            .AddItem "Form12.f12_131", 399
            .AddItem "Form12.f12_14", 400
            .AddItem "Form12.f12_141", 401
            .AddItem "Form12.f12_nlc", 402
            .AddItem "Form12.f12_f_1", 403
            .AddItem "Form12.f12_i_1", 404
            .AddItem "Form12.f12_o_1", 405
            .AddItem "Form12.f12_dr_1", 406
            .AddItem "Form12.f12_331_1", 407
            .AddItem "Form12.f12_9_1", 408
            .AddItem "Form12.f12_901_1", 409
            .AddItem "Form12.f12_10_1", 410
            .AddItem "Form12.f12_11_1", 411
            .AddItem "Form12.f12_12d1", 412
            .AddItem "Form12.f12_nlc_2", 413
            .AddItem "Form12.f12_f_2", 414
            .AddItem "Form12.f12_i_2", 415
            .AddItem "Form12.f12_o_2", 416
            .AddItem "Form12.f12_dr_2", 417
            .AddItem "Form12.f12_331_2", 418
            .AddItem "Form12.f12_9_2", 419
            .AddItem "Form12.f12_901_2", 420
            .AddItem "Form12.f12_10_2", 421
            .AddItem "Form12.f12_11_2", 422
            .AddItem "Form12.f12_12d2", 423
            .AddItem "Form12.f12_nlc_3", 424
            .AddItem "Form12.f12_f_3", 425
            .AddItem "Form12.f12_i_3", 426
            .AddItem "Form12.f12_o_3", 427
            .AddItem "Form12.f12_dr_3", 428
            .AddItem "Form12.f12_331_3", 429
            .AddItem "Form12.f12_9_3", 430
            .AddItem "Form12.f12_901_3", 431
            .AddItem "Form12.f12_10_3", 432
            .AddItem "Form12.f12_11_3", 433
            .AddItem "Form12.f12_12d3", 434
            .AddItem "Form12.f12_nlc_4", 435
            .AddItem "Form12.f12_f_4", 436
            .AddItem "Form12.f12_i_4", 437
            .AddItem "Form12.f12_o_4", 438
            .AddItem "Form12.f12_dr_4", 439
            .AddItem "Form12.f12_331_4", 440
            .AddItem "Form12.f12_9_4", 441
            .AddItem "Form12.f12_901_4", 442
            .AddItem "Form12.f12_10_4", 443
            .AddItem "Form12.f12_11_4", 444
            .AddItem "Form12.f12_12d4", 445
            .AddItem "Form12.f12_nlc_5", 446
            .AddItem "Form12.f12_f_5", 447
            .AddItem "Form12.f12_i_5", 448
            .AddItem "Form12.f12_o_5", 449
            .AddItem "Form12.f12_dr_5", 450
            .AddItem "Form12.f12_331_5", 451
            .AddItem "Form12.f12_9_5", 452
            .AddItem "Form12.f12_901_5", 453
            .AddItem "Form12.f12_10_5", 454
            .AddItem "Form12.f12_11_5", 455
            .AddItem "Form12.f12_12d5", 456
            .AddItem "Form12.f12_nlc_6", 457
            .AddItem "Form12.f12_f_6", 458
            .AddItem "Form12.f12_i_6", 459
            .AddItem "Form12.f12_o_6", 460
            .AddItem "Form12.f12_dr_6", 461
            .AddItem "Form12.f12_331_6", 462
            .AddItem "Form12.f12_9_6", 463
            .AddItem "Form12.f12_901_6", 464
            .AddItem "Form12.f12_10_6", 465
            .AddItem "Form12.f12_11_6", 466
            .AddItem "Form12.f12_12d6", 467
            .AddItem "Form12.f12_nlc_7", 468
            .AddItem "Form12.f12_f_7", 469
            .AddItem "Form12.f12_i_7", 470
            .AddItem "Form12.f12_o_7", 471
            .AddItem "Form12.f12_dr_7", 472
            .AddItem "Form12.f12_331_7", 473
            .AddItem "Form12.f12_9_7", 474
            .AddItem "Form12.f12_901_7", 475
            .AddItem "Form12.f12_10_7", 476
            .AddItem "Form12.f12_11_7", 477
            .AddItem "Form12.f12_12d7", 478
            .AddItem "Form12.f12_nlc_8", 479
            .AddItem "Form12.f12_f_8", 480
            .AddItem "Form12.f12_i_8", 481
            .AddItem "Form12.f12_o_8*", 482
            .AddItem "Form12.f12_dr_8", 483
            .AddItem "Form12.f12_331_8", 484
            .AddItem "Form12.f12_9_8", 485
            .AddItem "Form12.f12_901_8", 486
            .AddItem "Form12.f12_10_8", 487
            .AddItem "Form12.f12_11_8", 488
            .AddItem "Form12.f12_12d8", 489
            .AddItem "Form3.f3_n", 490
            .AddItem "Form3.f3_d", 491
            .AddItem "Form3.f3_1v", 492
            .AddItem "Form3.f3_2", 493
            .AddItem "Form3.f3_5kod", 494
            .AddItem "Form3.f3_5d", 495
            .AddItem "Form3.f3_6", 496
            .AddItem "Form3.f3_8", 497
            .AddItem "Form3.f3_8num", 498
            .AddItem "Form3.f3_8d", 499
            .AddItem "Form3.f3_9num", 500
            .AddItem "Form3.f3_9", 501
            .AddItem "Form3.f3_13", 502
            .AddItem "Form3.f3_14", 503
            .AddItem "Form3.f3_f", 504
            .AddItem "Form3.f3_i", 505
            .AddItem "Form3.f3_o", 506
            .AddItem "Form3.f3_15", 507
            .AddItem "Form3.f3_151", 508
            .AddItem "Form3.f3_18", 509
            .AddItem "Form3.f3_16", 510
            .AddItem "Form3.f3_161", 511
            .AddItem "Form3.f3_17_1", 512
            .AddItem "Form3.f3_17_2", 513
            .AddItem "Form3.f3_171", 514
            .AddItem "Form4.f4_n", 515
            .AddItem "Form4.f4_d", 516
            .AddItem "Form4.f4_1v", 517
            .AddItem "Form4.f4_3v", 518
            .AddItem "Form4.f4_4_1", 519
            .AddItem "Form4.f4_5kod", 520
            .AddItem "Form4.f4_5d", 521
            .AddItem "Form4.f4_6s", 522
            .AddItem "Form4.f4_6z", 523
            .AddItem "Form4.f4_6ch", 524
            .AddItem "Form4.f4_6p1_1", 525
            .AddItem "Form4.f4_6p1_2", 526
            .AddItem "Form4.f4_6p1_3", 527
            .AddItem "Form4.f4_6p1_4", 528
            .AddItem "Form4.f4_6p1_5", 529
            .AddItem "Form4.f4_7s", 530
            .AddItem "Form4.f4_7z", 531
            .AddItem "Form4.f4_7ch", 532
            .AddItem "Form4.f4_7p1_1", 533
            .AddItem "Form4.f4_7p1_2", 534
            .AddItem "Form4.f4_81", 535
            .AddItem "Form4.f4_8", 536
            .AddItem "Form4.f4_9_1", 537
            .AddItem "Form4.f4_9_2", 538
            .AddItem "Form4.f4_9_3", 539
            .AddItem "Form4.f4_9_4", 540
            .AddItem "Form4.f4_10", 541
            .AddItem "Form4.f4_101", 542
            .AddItem "Form4.f4_102", 543
            .AddItem "Form4.f4_103", 544
            .AddItem "Form4.f4_104", 545
            .AddItem "Form4.f4_105", 546
            .AddItem "Form4.f4_106", 547
            .AddItem "Form4.f4_107", 548
            .AddItem "Form4.f4_11", 549
            .AddItem "Form4.f4_111", 550
            .AddItem "Form4.f4_112", 551
            .AddItem "Form4.f4_113", 552
            .AddItem "Form4.f4_114", 553
            .AddItem "Form4.f4_115", 554
            .AddItem "Form4.f4_12", 555
            .AddItem "Form4.f4_121", 556
            .AddItem "Form4.f4_13", 557
            .AddItem "Form4.f4_15", 558
            .AddItem "Form4.f4_16", 559
            .AddItem "Form4.f4_161", 560
            .AddItem "Form4.f4_162", 561
            .AddItem "Form4.f4_163", 562
            .AddItem "Form4.f4_164", 563
            .AddItem "Form4.f4_165", 564
            .AddItem "Form4.f4_166", 565
            .AddItem "Form4.f4_167", 566
            .AddItem "Form4.f4_17", 567
            .AddItem "Form4.f4_171", 568
            .AddItem "Form4.f4_172", 569
            .AddItem "Form4.f4_173", 570
            .AddItem "Form4.f4_18", 571
            .AddItem "Form4.f4_181", 572
            .AddItem "Form4.f4_182", 573
            .AddItem "Form4.f4_183", 574
            .AddItem "Form4.f4_19", 575
            .AddItem "Form4.f4_191", 576
            .AddItem "Form4.f4_196", 577
            .AddItem "Form4.f4_192", 578
            .AddItem "Form4.f4_197", 579
            .AddItem "Form4.f4_193", 580
            .AddItem "Form4.f4_198", 581
            .AddItem "Form4.f4_194", 582
            .AddItem "Form4.f4_199", 583
            .AddItem "Form4.f4_195", 584
            .AddItem "Form4.f4_200", 585
            .AddItem "Form4.f4_20", 586
            .AddItem "Form4.f4_201", 587
            .AddItem "Form4.f4_21", 588
            .AddItem "Form4.f4_211", 589
            .AddItem "Form4.f4_212", 590
            .AddItem "Form4.f4_213", 591
            .AddItem "Form4.f4_214", 592
            .AddItem "Form4.f4_215", 593
            .AddItem "Form4.f4_216", 594
            .AddItem "Form4.f4_217", 595
            .AddItem "Form4.f4_218", 596
            .AddItem "Form4.f4_227", 597
            .AddItem "Form4.f4_22", 598
            .AddItem "Form4.f4_221", 599
            .AddItem "Form4.f4_222", 600
            .AddItem "Form4.f4_223", 601
            .AddItem "Form4.f4_224", 602
            .AddItem "Form4.f4_225", 603
            .AddItem "Form4.f4_228", 604
            .AddItem "Form4.f4_229", 605
            .AddItem "Form4.f4_241", 606
            .AddItem "Form4.f4_23", 607
            .AddItem "Form4.f4_231", 608
            .AddItem "Form4.f4_232", 609
            .AddItem "Form4.f4_233", 610
            .AddItem "Form4.f4_234", 611
            .AddItem "Form4.f4_237", 612
            .AddItem "Form4.f4_235", 613
            .AddItem "Form4.f4_238", 614
            .AddItem "Form4.f4_236", 615
            .AddItem "Form4.f4_239", 616
            .AddItem "Form4.f4_242", 617
            .AddItem "Form4.f4_249", 618
            .AddItem "Form4.f4_250", 619
            .AddItem "Form4.f4_243", 620
            .AddItem "Form4.f4_244", 621
            .AddItem "Form4.f4_245", 622
            .AddItem "Form4.f4_246", 623
            .AddItem "Form4.f4_247", 624
            .AddItem "Form4.f4_248", 625
            .AddItem "Form4.f4_25", 626
            .AddItem "Form4.f4_281", 627
            .AddItem "Form4.f4_282", 628
            .AddItem "Form4.f4_299", 629
            .AddItem "Form4.f4_29", 630
            .AddItem "Form4.f4_291", 631
            .AddItem "Form4.f4_292", 632
            .AddItem "Form4.f4_293", 633
            .AddItem "Form4.f4_294", 634
            .AddItem "Form4.f4_295", 635
            .AddItem "Form4.f4_29_1", 636
            .AddItem "Form4.f4_29_11", 637
            .AddItem "Form4.f4_296", 638
            .AddItem "Form4.f4_29_2", 639
            .AddItem "Form4.f4_29_21", 640
            .AddItem "Form4.f4_297", 641
            .AddItem "Form4.f4_30", 642
            .AddItem "Form4.f4_301", 643
            .AddItem "Form4.f4_302", 644
            .AddItem "Form4.f4_303", 645
            .AddItem "Form4.f4_304", 646
            .AddItem "Form4.f4_305", 647
            .AddItem "Form4.f4_306", 648
            .AddItem "Form4.f4_307", 649
            .AddItem "Form4.f4_308", 650
            .AddItem "Form4.f4_31", 651
            .AddItem "Form4.f4_311", 652
            .AddItem "Form4.f4_312", 653
            .AddItem "Form4.f4_313", 654
            .AddItem "Form4.f4_314", 655
            .AddItem "Form4.f4_315", 656
            .AddItem "Form4.f4_316", 657
            .AddItem "Form4.f4_317", 658
            .AddItem "Form4.f4_318", 659
            .AddItem "Form4.f4_32", 660
            .AddItem "Form4.f4_33", 661
            .AddItem "Form4.f4_331", 662
            .AddItem "Form4.f10_1", 663
            .AddItem "Form4.f10_2", 664
            .AddItem "Form4.f10_3", 665
            .AddItem "Form4.f10_4", 666
            .AddItem "Form2.f2_2kod", 667
            .AddItem "Form2.f2_3num", 668
            .AddItem "Form2.f2_4num", 669
            .AddItem "Form2.status", 670
            .AddItem "Form2.f2_1v", 671
            .AddItem "Form2.f2_2", 672
            .AddItem "Form2.f2_3v", 673
            .AddItem "Form2.f2_6kod", 674
            .AddItem "Form2.f2_7d", 675
            .AddItem "Form2.f2_fam", 676
            .AddItem "Form2.f2_imj", 677
            .AddItem "Form2.f2_otc", 678
            .AddItem "Form2.f2_10", 679
            .AddItem "Form2.f2_11", 680
            .AddItem "Form2.f2_11d", 681
            .AddItem "Form2.f2_111g", 682
            .AddItem "Form2.f2_111r", 683
            .AddItem "Form2.f2_111k", 684
            .AddItem "Form2.f2_12", 685
            .AddItem "Form2.f2_13_1", 686
            .AddItem "Form2.f2_13_2", 687
            .AddItem "Form2.f2_131", 688
            .AddItem "Form2.f2_132_1", 689
            .AddItem "Form2.f2_132_2", 690
            .AddItem "Form2.f2_14", 691
            .AddItem "Form2.f2_15v_1", 692
            .AddItem "Form2.f2_15v_2", 693
            .AddItem "Form2.f2_15kod", 694
            .AddItem "Form2.f2_152", 695
            .AddItem "Form2.f2_16", 696
            .AddItem "Form2.f2_17", 697
            .AddItem "Form2.f2_181", 698
            .AddItem "Form2.f2_182", 699
            .AddItem "Form2.f2_19", 700
            .AddItem "Form2.f2_193", 701
            .AddItem "Form2.f2_194", 702
            .AddItem "Form2.f2_191", 703
            .AddItem "Form2.f2_192", 704
            .AddItem "Form2.f2_20", 705
            .AddItem "Form2.f2_21s", 706
            .AddItem "Form2.f2_21z", 707
            .AddItem "Form2.f2_21ch", 708
            .AddItem "Form2.f2_21p1_1", 709
            .AddItem "Form2.f2_21p1_2", 710
            .AddItem "Form2.f2_21p1_3", 711
            .AddItem "Form2.f2_21p1_4", 712
            .AddItem "Form2.f2_21p1_5", 713
            .AddItem "Form2.f2_211s", 714
            .AddItem "Form2.f2_211z", 715
            .AddItem "Form2.f2_211ch", 716
            .AddItem "Form2.f2_211p1_1", 717
            .AddItem "Form2.f2_211p1_2", 718
            .AddItem "Form2.f2_211p1_3", 719
            .AddItem "Form2.f2_211p1_4", 720
            .AddItem "Form2.f2_211p1_5", 721
            .AddItem "Form2.f2_212s", 722
            .AddItem "Form2.f2_212z", 723
            .AddItem "Form2.f2_212ch", 724
            .AddItem "Form2.f2_212p1_1", 725
            .AddItem "Form2.f2_212p1_2", 726
            .AddItem "Form2.f2_212p1_3", 727
            .AddItem "Form2.f2_212p1_4", 728
            .AddItem "Form2.f2_212p1_5", 729
            .AddItem "Form2.f2_213s", 730
            .AddItem "Form2.f2_213z", 731
            .AddItem "Form2.f2_213ch", 732
            .AddItem "Form2.f2_213p1_1", 733
            .AddItem "Form2.f2_213p1_2", 734
            .AddItem "Form2.f2_213p1_3", 735
            .AddItem "Form2.f2_213p1_4", 736
            .AddItem "Form2.f2_213p1_5", 737
            .AddItem "Form2.f2_22s", 738
            .AddItem "Form2.f2_22z", 739
            .AddItem "Form2.f2_22ch", 740
            .AddItem "Form2.f2_22p1_1", 741
            .AddItem "Form2.f2_22p1_2", 742
            .AddItem "Form2.f2_23", 743
            .AddItem "Form2.f2_24", 744
            .AddItem "Form2.f2_25", 745
            .AddItem "Form2.f2_261", 746
            .AddItem "Form2.f2_26", 747
            .AddItem "Form2.f2_27_1", 748
            .AddItem "Form2.f2_27_2", 749
            .AddItem "Form2.f2_282", 750
            .AddItem "Form2.f2_28", 751
            .AddItem "Form2.f2_281", 752
            .AddItem "Form2.f2_29", 753
            .AddItem "Form2.f2_30", 754
            .AddItem "Form2.f2_31_1", 755
            .AddItem "Form2.f2_31_2", 756
            .AddItem "Form2.f2_31_3", 757
            .AddItem "Form2.f2_311_1", 758
            .AddItem "Form2.f2_311_2", 759
            .AddItem "Form2.f2_311_3", 760
            .AddItem "Form2.f2_32_1", 761
            .AddItem "Form2.f2_32_2", 762
            .AddItem "Form2.f2_32_3", 763
            .AddItem "Form2.f2_33_1", 764
            .AddItem "Form2.f2_33_2", 765
            .AddItem "Form2.f2_33_3", 766
            .AddItem "Form2.f2_33_4", 767
            .AddItem "Form2.f2_34_1", 768
            .AddItem "Form2.f2_34_2", 769
            .AddItem "Form2.f2_34_3", 770
            .AddItem "Form2.f2_34_4", 771
            .AddItem "Form2.f2_35_1", 772
            .AddItem "Form2.f2_35_2", 773
            .AddItem "Form2.f2_36_1", 774
            .AddItem "Form2.f2_36_2", 775
            .AddItem "Form2.f2_38", 776
            .AddItem "Form2.f2_381", 777
            .AddItem "Form2.f2_382", 778
            .AddItem "Form2.f2_383", 779
            .AddItem "Form2.f2_39", 780
            .AddItem "Form2.f2_391", 781
            .AddItem "Form2.f2_40", 782
            .AddItem "Form2.f2_41", 783
            .AddItem "Form2.f2_411_1", 784
            .AddItem "Form2.f2_411_2", 785
            .AddItem "Form2.f2_42_1", 786
            .AddItem "Form2.f2_42_2", 787
            .AddItem "Form2.f2_43", 788
            .AddItem "Form2.f2_44", 789
            .AddItem "Form2.f2_45_1", 790
            .AddItem "Form2.f2_45_2", 791
            .AddItem "Form2.f2_46", 792
            .AddItem "Form2.f2_47", 793
            .AddItem "Form2.f2_48", 794
            .AddItem "Form2.f2_49", 795
            .AddItem "Form2.f2_50_1", 796
            .AddItem "Form2.f2_50_2", 797
            .AddItem "Form2.f2_50_#", 798
            .AddItem "Form2.f2_51kod", 799
            .AddItem "Form2.f2_51d", 800
            .AddItem "Form2.f2_52", 801
            .AddItem "Form2.f2_53", 802
            .AddItem "Form2.f2_54", 803
            .AddItem "Form2.f2_55_1", 804
            .AddItem "Form2.f2_55_2", 805
            .AddItem "Form2.f2_56", 806
            .AddItem "Form2.f2_56d", 807
            .AddItem "Form2.f2_60_1", 808
            .AddItem "Form2.f2_60_2", 809
            .AddItem "Form2.f2_60_3", 810
            .AddItem "Form2.f2_61_1", 811
            .AddItem "Form2.f2_61_2", 812
            .AddItem "Form2.usr_f6", 813
            .AddItem "Form2.dat_f6", 814
            .AddItem "Form2.cur", 815
            .AddItem "Form2.f2_n", 816
            .AddItem "Form2.f2_d", 817
            .AddItem "Form6.f6_1kod", 818
            .AddItem "Form6.f6_1", 819
            .AddItem "Form6.f6_3", 820
            .AddItem "Form6.f6_4", 821
            .AddItem "Form6.f6_5", 822
            .AddItem "Form6.f6_6", 823
            .AddItem "Form6.f6_7", 824
            .AddItem "Form6.f6_7_1", 825
            .AddItem "Form6.f6_7_2", 826
            .AddItem "Form6.f6_8", 827
            .AddItem "Form6.f6_9", 828
            .AddItem "Form6.f6_10", 829
            .AddItem "Form6.f6_11", 830
            .AddItem "Form6.f6_12", 831
            .AddItem "Form6.f6_12_1", 832
            .AddItem "Form6.f6_12_2", 833
            .AddItem "Form6.f6_13", 834
            .AddItem "Form6.f6_14", 835
            .AddItem "Form6.f6_151v1", 836
            .AddItem "Form6.f6_151v2", 837
            .AddItem "Form6.f6_151s", 838
            .AddItem "Form6.f6_151z", 839
            .AddItem "Form6.f6_151ch", 840
            .AddItem "Form6.f6_151p1_1", 841
            .AddItem "Form6.f6_151p1_2", 842
            .AddItem "Form6.f6_151p1_3", 843
            .AddItem "Form6.f6_151p1_4", 844
            .AddItem "Form6.f6_151p1_5", 845
            .AddItem "Form6.f6_152v1", 846
            .AddItem "Form6.f6_152v2", 847
            .AddItem "Form6.f6_152s", 848
            .AddItem "Form6.f6_152z", 849
            .AddItem "Form6.f6_152ch", 850
            .AddItem "Form6.f6_152p1_1", 851
            .AddItem "Form6.f6_152p1_2", 852
            .AddItem "Form6.f6_152p1_3", 853
            .AddItem "Form6.f6_152p1_4", 854
            .AddItem "Form6.f6_152p1_5", 855
            .AddItem "Form6.f6_153v1", 856
            .AddItem "Form6.f6_153v2", 857
            .AddItem "Form6.f6_153s", 858
            .AddItem "Form6.f6_153z", 859
            .AddItem "Form6.f6_153ch", 860
            .AddItem "Form6.f6_153p1_1", 861
            .AddItem "Form6.f6_153p1_2", 862
            .AddItem "Form6.f6_153p1_3", 863
            .AddItem "Form6.f6_153p1_4", 864
            .AddItem "Form6.f6_153p1_5", 865
            .AddItem "Form6.f6_154v1", 866
            .AddItem "Form6.f6_154v2", 867
            .AddItem "Form6.f6_154s", 868
            .AddItem "Form6.f6_154z", 869
            .AddItem "Form6.f6_154ch", 870
            .AddItem "Form6.f6_154p1_1", 871
            .AddItem "Form6.f6_154p1_2", 872
            .AddItem "Form6.f6_154p1_3", 873
            .AddItem "Form6.f6_154p1_4", 874
            .AddItem "Form6.f6_154p1_5", 875
            .AddItem "Form6.f6_155v1", 876
            .AddItem "Form6.f6_155v2", 877
            .AddItem "Form6.f6_155s", 878
            .AddItem "Form6.f6_155z", 879
            .AddItem "Form6.f6_155ch", 880
            .AddItem "Form6.f6_155p1_1", 881
            .AddItem "Form6.f6_155p1_2", 882
            .AddItem "Form6.f6_155p1_3", 883
            .AddItem "Form6.f6_155p1_4", 884
            .AddItem "Form6.f6_155p1_5", 885
            .AddItem "Form6.f6_156v1", 886
            .AddItem "Form6.f6_156v2", 887
            .AddItem "Form6.f6_156s", 888
            .AddItem "Form6.f6_156z", 889
            .AddItem "Form6.f6_156ch", 890
            .AddItem "Form6.f6_156p1_1", 891
            .AddItem "Form6.f6_156p1_2", 892
            .AddItem "Form6.f6_156p1_3", 893
            .AddItem "Form6.f6_156p1_4", 894
            .AddItem "Form6.f6_156p1_5", 895
            .AddItem "Form6.f6_161", 896
            .AddItem "Form6.f6_162", 897
            .AddItem "Form6.f6_163", 898
            .AddItem "Form6.f6_164", 899
            .AddItem "Form6.f6_165", 900
            .AddItem "Form6.f6_166", 901
            .AddItem "Form6.f6_171v1", 902
            .AddItem "Form6.f6_171v2", 903
            .AddItem "Form6.f6_171s", 904
            .AddItem "Form6.f6_171z", 905
            .AddItem "Form6.f6_171ch", 906
            .AddItem "Form6.f6_171p1_1", 907
            .AddItem "Form6.f6_171p1_2", 908
            .AddItem "Form6.f6_171p1_3", 909
            .AddItem "Form6.f6_171p1_4", 910
            .AddItem "Form6.f6_171p1_5", 911
            .AddItem "Form6.f6_172v1", 912
            .AddItem "Form6.f6_172v2", 913
            .AddItem "Form6.f6_172s", 914
            .AddItem "Form6.f6_172z", 915
            .AddItem "Form6.f6_172ch", 916
            .AddItem "Form6.f6_172p1_1", 917
            .AddItem "Form6.f6_172p1_2", 918
            .AddItem "Form6.f6_172p1_3", 919
            .AddItem "Form6.f6_172p1_4", 920
            .AddItem "Form6.f6_172p1_5", 921
            .AddItem "Form6.f6_173v1", 922
            .AddItem "Form6.f6_173v2", 923
            .AddItem "Form6.f6_173s", 924
            .AddItem "Form6.f6_173z", 925
            .AddItem "Form6.f6_173ch", 926
            .AddItem "Form6.f6_173p1_1", 927
            .AddItem "Form6.f6_173p1_2", 928
            .AddItem "Form6.f6_173p1_3", 929
            .AddItem "Form6.f6_173p1_4", 930
            .AddItem "Form6.f6_173p1_5", 931
            .AddItem "Form6.f6_174v1", 932
            .AddItem "Form6.f6_174v2", 933
            .AddItem "Form6.f6_174s", 934
            .AddItem "Form6.f6_174z", 935
            .AddItem "Form6.f6_174ch", 936
            .AddItem "Form6.f6_174p1_1", 937
            .AddItem "Form6.f6_174p1_2", 938
            .AddItem "Form6.f6_174p1_3", 939
            .AddItem "Form6.f6_174p1_4", 940
            .AddItem "Form6.f6_174p1_5", 941
            .AddItem "Form6.f6_175v1", 942
            .AddItem "Form6.f6_175v2", 943
            .AddItem "Form6.f6_175s", 944
            .AddItem "Form6.f6_175z", 945
            .AddItem "Form6.f6_175ch", 946
            .AddItem "Form6.f6_175p1_1", 947
            .AddItem "Form6.f6_175p1_2", 948
            .AddItem "Form6.f6_175p1_3", 949
            .AddItem "Form6.f6_175p1_4", 950
            .AddItem "Form6.f6_175p1_5", 951
            .AddItem "Form6.f6_176v1", 952
            .AddItem "Form6.f6_176v2", 953
            .AddItem "Form6.f6_176s", 954
            .AddItem "Form6.f6_176z", 955
            .AddItem "Form6.f6_176ch", 956
            .AddItem "Form6.f6_176p1_1", 957
            .AddItem "Form6.f6_176p1_2", 958
            .AddItem "Form6.f6_176p1_3", 959
            .AddItem "Form6.f6_176p1_4", 960
            .AddItem "Form6.f6_176p1_5", 961
            .AddItem "Form6.f6_181v1", 962
            .AddItem "Form6.f6_181l1", 963
            .AddItem "Form6.f6_181m1", 964
            .AddItem "Form6.f6_181d1", 965
            .AddItem "Form6.f6_181h1", 966
            .AddItem "Form6.f6_181v2", 967
            .AddItem "Form6.f6_181l2", 968
            .AddItem "Form6.f6_181m2", 969
            .AddItem "Form6.f6_181d2", 970
            .AddItem "Form6.f6_181h2", 971
            .AddItem "Form6.f6_182v1", 972
            .AddItem "Form6.f6_182l1", 973
            .AddItem "Form6.f6_182m1", 974
            .AddItem "Form6.f6_182d1", 975
            .AddItem "Form6.f6_182v2", 976
            .AddItem "Form6.f6_182l2", 977
            .AddItem "Form6.f6_182m2", 978
            .AddItem "Form6.f6_182d2", 979
            .AddItem "Form6.f6_18d", 980
            .AddItem "Form6.f6_191v1", 981
            .AddItem "Form6.f6_191l1", 982
            .AddItem "Form6.f6_191m1", 983
            .AddItem "Form6.f6_191d1", 984
            .AddItem "Form6.f6_191h1", 985
            .AddItem "Form6.f6_191s1", 986
            .AddItem "Form6.f6_191v2", 987
            .AddItem "Form6.f6_191l2", 988
            .AddItem "Form6.f6_191m2", 989
            .AddItem "Form6.f6_191d2", 990
            .AddItem "Form6.f6_191h2", 991
            .AddItem "Form6.f6_191s2", 992
            .AddItem "Form6.f6_192v1", 993
            .AddItem "Form6.f6_192l1", 994
            .AddItem "Form6.f6_192m1", 995
            .AddItem "Form6.f6_192d1", 996
            .AddItem "Form6.f6_192s1", 997
            .AddItem "Form6.f6_192v2", 998
            .AddItem "Form6.f6_192l2", 999
            .AddItem "Form6.f6_192m2", 1000
            .AddItem "Form6.f6_192d2", 1001
            .AddItem "Form6.f6_192s2", 1002
            .AddItem "Form6.f6_19_3_1", 1003
            .AddItem "Form6.f6_19_3_2", 1004
            .AddItem "Form6.f6_19_3l", 1005
            .AddItem "Form6.f6_19_3m", 1006
            .AddItem "Form6.f6_19_4", 1007
            .AddItem "Form6.f6_19_5", 1008
            .AddItem "Form6.f6_20", 1009
            .AddItem "Form6.f6_n", 1010
            .AddItem "Form6.f6_d", 1011
            .AddItem "Form6.f6_19_6", 1012
            .AddItem "Form2.f2_541", 1013
            .AddItem "Form2.f2_561", 1014
            .AddItem "Form2.f2_4d", 1015
            .AddItem "Form2.f2_36_1_1", 1016
            .AddItem "Form2.f2_36_1_2", 1017
            .AddItem "Form2.f2_36_1_3", 1018
            .AddItem "Form2.f2_36_1_4", 1019
            .AddItem "Form2.f2_36_1_5", 1020
            .AddItem "Form2.f2_28_1", 1021
            .AddItem "Form2.f2_511k", 1022
            .AddItem "Form2.f2_511d", 1023
            .AddItem "Form2.f2_542", 1024
            .AddItem "Form5.f5_1kod", 1025
            .AddItem "Form5.f5_3num", 1026
            .AddItem "Form5.f5_4num", 1027
            .AddItem "Form5.f5_1v", 1028
            .AddItem "Form5.f5_2", 1029
            .AddItem "Form5.f5_3v", 1030
            .AddItem "Form5.f5_6kod", 1031
            .AddItem "Form5.f5_6d", 1032
            .AddItem "Form5.f5_7s", 1033
            .AddItem "Form5.f5_7z", 1034
            .AddItem "Form5.f5_7ch", 1035
            .AddItem "Form5.f5_7p1_1", 1036
            .AddItem "Form5.f5_7p1_2", 1037
            .AddItem "Form5.f5_7p1_3", 1038
            .AddItem "Form5.f5_7p1_4", 1039
            .AddItem "Form5.f5_7p1_5", 1040
            .AddItem "Form5.f5_701s", 1041
            .AddItem "Form5.f5_701z", 1042
            .AddItem "Form5.f5_701ch", 1043
            .AddItem "Form5.f5_701p1_1", 1044
            .AddItem "Form5.f5_701p1_2", 1045
            .AddItem "Form5.f5_701p1_3", 1046
            .AddItem "Form5.f5_701p1_4", 1047
            .AddItem "Form5.f5_701p1_5", 1048
            .AddItem "Form5.f5_702s", 1049
            .AddItem "Form5.f5_702z", 1050
            .AddItem "Form5.f5_702ch", 1051
            .AddItem "Form5.f5_702p1_1", 1052
            .AddItem "Form5.f5_702p1_2", 1053
            .AddItem "Form5.f5_702p1_3", 1054
            .AddItem "Form5.f5_702p1_4", 1055
            .AddItem "Form5.f5_702p1_5", 1056
            .AddItem "Form5.f5_703s", 1057
            .AddItem "Form5.f5_703z", 1058
            .AddItem "Form5.f5_703ch", 1059
            .AddItem "Form5.f5_703p1_1", 1060
            .AddItem "Form5.f5_703p1_2", 1061
            .AddItem "Form5.f5_703p1_3", 1062
            .AddItem "Form5.f5_703p1_4", 1063
            .AddItem "Form5.f5_703p1_5", 1064
            .AddItem "Form5.f5_8", 1065
            .AddItem "Form5.f5_fam", 1066
            .AddItem "Form5.f5_imj", 1067
            .AddItem "Form5.f5_otc", 1068
            .AddItem "Form5.f5_9", 1069
            .AddItem "Form5.f5_101d", 1070
            .AddItem "Form5.f5_102", 1071
            .AddItem "Form5.f5_103g", 1072
            .AddItem "Form5.f5_103r", 1073
            .AddItem "Form5.f5_103k", 1074
            .AddItem "Form5.f5_104", 1075
            .AddItem "Form5.f5_105", 1076
            .AddItem "Form5.f5_106_1", 1077
            .AddItem "Form5.f5_106_2", 1078
            .AddItem "Form5.f5_107_1", 1079
            .AddItem "Form5.f5_107_2", 1080
            .AddItem "Form5.f5_11v", 1081
            .AddItem "Form5.f5_11", 1082
            .AddItem "Form5.f5_121", 1083
            .AddItem "Form5.f5_122", 1084
            .AddItem "Form5.f5_13", 1085
            .AddItem "Form5.f5_131", 1086
            .AddItem "Form5.f5_132", 1087
            .AddItem "Form5.f5_133", 1088
            .AddItem "Form5.f5_134", 1089
            .AddItem "Form5.f5_14", 1090
            .AddItem "Form5.f5_141", 1091
            .AddItem "Form5.f5_15", 1092
            .AddItem "Form5.f5_16", 1093
            .AddItem "Form5.f5_171", 1094
            .AddItem "Form5.f5_172", 1095
            .AddItem "Form5.f5_18", 1096
            .AddItem "Form5.f5_1811", 1097
            .AddItem "Form5.f5_181", 1098
            .AddItem "Form5.f5_1821", 1099
            .AddItem "Form5.f5_1822", 1100
            .AddItem "Form5.f5_1823", 1101
            .AddItem "Form5.f5_183", 1102
            .AddItem "Form5.f5_19", 1103
            .AddItem "Form5.f5_20", 1104
            .AddItem "Form5.f5_21_1", 1105
            .AddItem "Form5.f5_21_2", 1106
            .AddItem "Form5.f5_21_3", 1107
            .AddItem "Form5.f5_21_4", 1108
            .AddItem "Form5.f5_31num", 1109
            .AddItem "Form5.f5_31d", 1110
            .AddItem "Form5.f5_n", 1111
            .AddItem "Form5.f5_d", 1112
            .AddItem "Form5.f5_4_1", 1113
            .AddItem "Form5.f5_18_4", 1114
            .AddItem "Fabula.Œ¬ƒ", 1115
            .AddItem "Fabula.ÕŒÃ≈– œ–≈—“", 1116
            .AddItem "Fabula.Œ—Õ", 1117
            .AddItem "Fabula.ƒ¿“¿ ¬Œ«¡”∆ƒ≈Õ»ﬂ", 1118
            .AddItem "Fabula.œŒ—“ »÷", 1119
            .AddItem "Fabula.—“", 1120
            .AddItem "Fabula.«", 1121
            .AddItem "Fabula.◊", 1122
            .AddItem "Fabula.œ1", 1123
            .AddItem "Fabula.œ2", 1124
            .AddItem "Fabula.œ3", 1125
            .AddItem "Fabula.ƒ¿“¿ –≈ÿ≈Õ»ﬂ", 1126
            .AddItem "Fabula.‘¿¡”À¿", 1127
            .AddItem "Fabula.God", 1128
            End If
            End If
            
            
    Case Else
          .SetFocus
          .RowSourceType = "Value List"
          .AddItem "", 0
          .AddItem "Form11", 1
          .AddItem "Form12", 2
          .AddItem "Form2", 3
          .AddItem "Form3", 4
          .AddItem "Form4", 5
          .AddItem "Form5", 6
          .AddItem "Form6", 7
          .AddItem "‘‡·ÛÎ˚", 8
    End Select
    End If
    
End With
Next ctl
Me.Form1Unselect.Enabled = False
Me.Spisok_Field1.Enabled = False
Me.Spisok_Field2.Enabled = False
End Sub
Private Sub Query_Click()
Dim ctl As Control
Dim fieldsname, formname As String
Dim qdf As QueryDef
Dim ssql As String
Dim sWhereClause As String
Dim strval As String
Dim intLastPos, intMaxLen As Integer
Dim Counter As Integer
Dim form1name, form2name, form12name, form3name, form4name, form5name, form6name, formelsename As String
Dim forma As String


sWhereClause = " WHERE "
fieldsname = "SELECT"
For Each ctl In Me.Controls
    With ctl
Select Case .ControlType
            Case acTextBox
                .SetFocus
               If sWhereClause = " WHERE " And Left(ctl.Name, 2) = "st" Then
                    If .Text <> "" Then
                        sWhereClause = sWhereClause & BuildCriteria("Form1.f1_13s", dbInteger, .Text)
                    End If
               Else
                    If .Text <> "" Then
                            If Left(ctl.Name, 2) = "st" Then
                                sWhereClause = sWhereClause & " OR " & BuildCriteria("Form1.f1_13s", dbInteger, .Text)
                            End If
                            If Left(ctl.Name, 2) = "zn" Then
                            sWhereClause = sWhereClause & " AND " & BuildCriteria("Form1.f1_13z", dbInteger, .Text)
                            End If
                            If Left(ctl.Name, 2) = "ch" Then
                            sWhereClause = sWhereClause & " AND " & BuildCriteria("Form1.f1_13ch", dbInteger, .Text)
                            End If
                            If Left(ctl.Name, 2) = "p1" Then
                            sWhereClause = sWhereClause & " AND " & BuildCriteria("Form1.f1_13p1_1", dbInteger, .Text)
                            End If
                            If Left(ctl.Name, 2) = "p2" Then
                            sWhereClause = sWhereClause & " AND " & BuildCriteria("Form1.f1_13p1_2", dbInteger, .Text)
                            End If
                    End If
               End If
               Case acCheckBox
                    If ctl.Value = True Then
                        sWhereClause = sWhereClause + "AND Form1.f1_7d >= 20160112 AND (Form1.f1_2 = 1 OR Form1.f1_2 = 3) "
                        sWhereClause = sWhereClause + "AND ((Form11.f11_25k NOT IN (12,13,14,15,16,17,18,20,21,25,26,29,32,33,34,36,37,38,40,43,44,46,47,48,49,50) AND Form11.f11_28 = 0) "
                        sWhereClause = sWhereClause + "OR (Form11.f11_25k IN (1,31,8,35,61) AND Form11.f11_6d >= 20160111 AND Form11.f11_28 not in (12,13,14,15,16,17,18,20,21,25,26,29,32,33,34,36,37,38,40,43,44,46,47,48,49,50,59)))"
                    End If
             Case acComboBox
                    If ctl.Value <> "" Then
               
                            If Left(ctl.Name, 4) = "form" Then
                                     Counter = Counter + 1
                                If Counter >= 3 Then
                                formname = "( " & formname & " )"
                                End If
                            '********************************************
                            If formname = "" Then
                                    formname = ctl.Value
                       
                           '*********************************************
                            Else
                                Select Case ctl.Value
                                Case "Form11"
                                    formname = formname + " LEFT JOIN " + ctl.Value + " ON (Form1.[f1_4num] = Form11.[f1_4num]) AND (Form1.[f1_3num]  = Form11.[f1_3num])"
                                Case "Form2"
                                    formname = formname + " LEFT JOIN " + ctl.Value + "  ON (Form1.[f1_4num] = Form2.[f2_4num])   AND (Form1.[f1_3num] = Form2.[f2_3num])"
                                Case "Form12"
                                    formname = formname + " LEFT JOIN " + ctl.Value + "  ON (Form1.[f1_4num] = Form12.[f1_4num]) AND (Form1.[f1_3num]  = Form12.[f1_3num])"
                                Case "Form3"
                                    formname = formname + "  LEFT JOIN " + ctl.Value + " ON (Form1.[f1_3num] = Form3.[f1_3num])   AND (Form1.[f1_4num] = Form3.[f1_4num])"
                                Case "Form4"
                                    formname = formname + "  LEFT JOIN " + ctl.Value + " ON (Form1.[f1_4num] = Form4.[f1_4num])   AND (Form1.[f1_3num] = Form4.[f1_3num])"
                                Case "Form5"
                                    formname = formname + "  LEFT JOIN " + ctl.Value + " ON (Form1.[f1_3num] = Form5.[f5_3num])"
                                Case "Form6"
                                    formname = formname + " LEFT JOIN " + ctl.Value + "  ON (Form1.[f1_4num] = Form6.[f2_4num])   AND (Form1.[f1_3num] = Form6.[f2_3num])"
                                Case Else
                                    formname = formname + " LEFT JOIN " + ctl.Value + "  ON (Form1.[f1_4num] = ‘‡·ÛÎ˚.[Œ—Õ])     AND (Form1.[f1_3num] = ‘‡·ÛÎ˚.[ÕŒÃ≈– œ–≈—“])"
                                End Select

                            End If
                        Else
                            If fieldsname = "SELECT" Then
                            fieldsname = fieldsname + " " + ctl.Value & ","
                            Else
                            fieldsname = fieldsname + " " + ctl.Value & ","
                            End If
                    
                    End If
                    End If
                    
                    
    End Select
    End With
 
    Next ctl
If sWhereClause = " WHERE " Then
Exit Sub
End If
formname = " FROM " & formname

intMaxLen = Len(fieldsname)
fieldsname = Left(fieldsname, intMaxLen - 1)
ssql = fieldsname + formname + sWhereClause
On Error Resume Next
DoCmd.DeleteObject acQuery, "tempQry"
On Error GoTo 0
Set qdf = CurrentDb.CreateQueryDef("tempQry", ssql)
DoCmd.Close acForm, Me.Name, acSaveNo
Dim fDialog As Object
Dim filePath As String
Set fDialog = Application.FileDialog(2)
With fDialog
.InitialFileName = "C:\«‡ÔÓÒ"
If .Show <> -1 Then
DoCmd.OpenForm "F_Index", acNormal
Exit Sub
End If
filePath = .InitialFileName
End With
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "tempQry", filePath
DoCmd.OpenForm "F_Index", acNormal
End Sub


