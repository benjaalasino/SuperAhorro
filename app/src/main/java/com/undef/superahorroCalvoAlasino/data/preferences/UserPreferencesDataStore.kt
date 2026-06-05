package com.undef.superahorroCalvoAlasino.data.preferences

import android.content.Context
import androidx.datastore.core.DataStore
import androidx.datastore.preferences.core.Preferences
import androidx.datastore.preferences.preferencesDataStore

val Context.userPrefs: DataStore<Preferences> by preferencesDataStore(name = "user_prefs")
