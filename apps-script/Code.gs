/* WatchTogether Google-hosted app. Deploy as USER_ACCESSING, require Google sign-in.
 * Private helpers end in _ so HtmlService cannot invoke them as public functions.
 * Run setupWatchTogether once after setting script properties. */
var SCHEMA_ = {
  users: [
    "id",
    "email",
    "google_sub",
    "display_name",
    "streaming_services",
    "region",
    "timezone",
  ],
  titles: [
    "id",
    "tmdb_id",
    "media_type",
    "title",
    "year",
    "runtime_min",
    "genres",
    "keywords",
    "overview",
    "poster_path",
    "cast_top",
    "director",
    "providers",
    "metadata_updated_at",
  ],
  episodes: ["key", "title_id", "season", "episode", "name", "air_date"],
  tracking: [
    "key",
    "user_id",
    "title_id",
    "status",
    "last_season",
    "last_episode",
    "tier",
    "rank_pos",
    "score",
    "seen_before",
    "follow",
    "revision",
    "updated_at",
  ],
  comparisons: ["id", "user_id", "title_a", "title_b", "winner", "compared_at"],
  watch_events: [
    "key",
    "user_id",
    "title_id",
    "season",
    "episode",
    "started_at",
    "completed",
    "source",
    "service",
  ],
  aliases: [
    "key",
    "raw_string",
    "service",
    "title_id",
    "confidence",
    "confirmed_by_user",
  ],
  import_reviews: [
    "id",
    "key",
    "user_id",
    "raw_string",
    "service",
    "watched_at",
    "season",
    "episode",
    "candidates",
    "resolved_at",
  ],
  operations: [
    "key",
    "user_id",
    "request_id",
    "fingerprint",
    "status",
    "result",
    "updated_at",
  ],
};
var JSON_COLUMNS_ = [
  "streaming_services",
  "genres",
  "keywords",
  "cast_top",
  "providers",
  "episode_data",
  "candidates",
  "result",
];

function properties_() {
  return PropertiesService.getScriptProperties();
}
function workbook_() {
  var id = properties_().getProperty("SPREADSHEET_ID");
  if (!id) throw new Error("Set SPREADSHEET_ID in Script Properties");
  return SpreadsheetApp.openById(id);
}
function locked_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000))
    throw new Error("BUSY: another edit is saving. Retry shortly.");
  try {
    return fn();
  } finally {
    try {
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  }
}
function setupWatchTogether() {
  allowedEmail_();
  return locked_(function () {
    var props = properties_(),
      book = workbook_();
    Object.keys(SCHEMA_).forEach(function (name) {
      var sheet = book.getSheetByName(name) || book.insertSheet(name);
      var columns = SCHEMA_[name];
      if (sheet.getLastRow() === 0) {
        sheet.getRange(1, 1, 1, columns.length).setValues([columns]);
        sheet.setFrozenRows(1);
      } else {
        var headers = sheet.getRange(1, 1, 1, columns.length).getValues()[0];
        if (JSON.stringify(headers) !== JSON.stringify(columns))
          throw new Error(
            "Unexpected " +
              name +
              " headers; existing data was not overwritten",
          );
      }
    });
    var accounts = [
      {
        id: "a",
        email: props.getProperty("USER_A_EMAIL"),
        display_name: props.getProperty("USER_A_NAME") || "Account A",
      },
      {
        id: "b",
        email: props.getProperty("USER_B_EMAIL"),
        display_name: props.getProperty("USER_B_NAME") || "Account B",
      },
    ];
    if (
      !accounts[0].email ||
      !accounts[1].email ||
      accounts[0].email.toLowerCase() === accounts[1].email.toLowerCase()
    )
      throw new Error(
        "Set two distinct USER_A_EMAIL and USER_B_EMAIL properties",
      );
    var existing = records_("users");
    if (existing.length > 2) throw new Error("Only two accounts are supported");
    accounts.forEach(function (user) {
      user.email = user.email.toLowerCase().trim();
      var old = existing.find(function (r) {
        return r.id === user.id;
      });
      if (old && old.email !== user.email)
        throw new Error(
          "An existing account has a different email; migrate it explicitly before changing configuration",
        );
      if (!old)
        put_(
          "users",
          user.id,
          Object.assign(user, {
            google_sub: "",
            streaming_services: [],
            region: props.getProperty("REGION") || "US",
            timezone: props.getProperty("TIMEZONE") || "America/New_York",
          }),
        );
    });

    return { ok: true, url: book.getUrl() };
  });
}
function records_(name) {
  var s = workbook_().getSheetByName(name);
  if (!s) throw new Error("Run setupWatchTogether first");
  if (s.getLastRow() < 2) return [];
  var columns = SCHEMA_[name];
  return s
    .getRange(2, 1, s.getLastRow() - 1, columns.length)
    .getValues()
    .map(function (values, i) {
      var obj = { _row: i + 2 };
      columns.forEach(function (k, j) {
        var v = values[j];
        if (v instanceof Date) v = v.toISOString();
        if (
          typeof v === "string" &&
          v.charAt(0) === "'" &&
          /^'[=+@\-']/.test(v)
        )
          v = v.slice(1);
        if (JSON_COLUMNS_.indexOf(k) >= 0) {
          try {
            v = v ? JSON.parse(v) : null;
          } catch (e) {
            throw new Error("Invalid JSON cell in " + name + "." + k);
          }
        }
        obj[k] = v === "" ? null : v;
      });
      return obj;
    });
}
function safeCell_(v, k) {
  if (v === undefined || v === null) return "";
  if (JSON_COLUMNS_.indexOf(k) >= 0) v = JSON.stringify(v);
  if (typeof v === "string") {
    if (v.length > 45000)
      throw new Error("Record exceeds the Google Sheets cell limit");
    if (/^[=+@\-']/.test(v)) v = "'" + v;
  }
  return v;
}
function put_(name, key, obj) {
  var s = workbook_().getSheetByName(name),
    columns = SCHEMA_[name],
    found = records_(name).find(function (r) {
      return String(r[columns[0]]) === String(key);
    });
  var values = columns.map(function (k) {
    return safeCell_(obj[k], k);
  });
  s.getRange(
    found ? found._row : s.getLastRow() + 1,
    1,
    1,
    columns.length,
  ).setValues([values]);
  return obj;
}
function remove_(name, key) {
  var columns = SCHEMA_[name],
    row = records_(name).find(function (r) {
      return String(r[columns[0]]) === String(key);
    });
  if (row) workbook_().getSheetByName(name).deleteRow(row._row);
}
function cleanRecord_(r) {
  if (!r) return null;
  var copy = Object.assign({}, r);
  delete copy._row;
  return copy;
}
function digest_(value) {
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    value,
    Utilities.Charset.UTF_8,
  )
    .map(function (b) {
      return ("0" + ((b + 256) % 256).toString(16)).slice(-2);
    })
    .join("");
}
function now_() {
  return new Date().toISOString();
}
function allowedEmail_() {
  var email = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  var props = properties_();
  var allowed = [props.getProperty("USER_A_EMAIL"), props.getProperty("USER_B_EMAIL")]
    .map(function (value) { return String(value || "").trim().toLowerCase(); });
  if (!email || allowed.indexOf(email) < 0)
    throw new Error("FORBIDDEN: sign in with one of the two configured Google accounts. Deploy this app as the user accessing it.");
  return email;
}
function identity_() {
  var email = allowedEmail_();
  var user = records_("users").find(function (row) { return row.email === email; });
  if (!user) throw new Error("Run setupWatchTogether first");
  return user;
}
// The only browser API. Google supplies the caller's identity, never the request.
function watchTogetherApi(request) {
  if (!request || typeof request !== "object" || Array.isArray(request) ||
      JSON.stringify(request).length > 180000) throw new Error("Invalid request");
  allowedEmail_();
  return locked_(function () {
    var user = identity_();
    return JSON.parse(JSON.stringify(dispatch_(user, request)));
  });
}
function doPost() {
  return output_({ error: "External POST access is disabled. Open the Google-hosted app." });
}
function doGet() {
  var template = HtmlService.createTemplateFromFile("Index");
  return template.evaluate().setTitle("WatchTogether")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}
function output_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
function episodesFor_(id) {
  return records_("episodes")
    .filter(function (e) {
      return Number(e.title_id) === Number(id);
    })
    .map(function (e) {
      return {
        season_number: e.season,
        episode_number: e.episode,
        name: e.name,
        air_date: e.air_date,
      };
    });
}
function attachEpisodes_(title) {
  return Object.assign(cleanRecord_(title), {
    episode_data: episodesFor_(title.id),
  });
}
function putEpisodes_(titleId, episodes) {
  var sheet = workbook_().getSheetByName("episodes"),
    existing = records_("episodes"),
    columns = SCHEMA_.episodes,
    additions = [];
  episodes.forEach(function (e) {
    var key = titleId + ":" + e.season_number + ":" + e.episode_number,
      row = {
        key: key,
        title_id: titleId,
        season: e.season_number,
        episode: e.episode_number,
        name: e.name,
        air_date: e.air_date,
      },
      old = existing.find(function (x) {
        return x.key === key;
      }),
      values = columns.map(function (k) {
        return safeCell_(row[k], k);
      });
    if (old) sheet.getRange(old._row, 1, 1, columns.length).setValues([values]);
    else additions.push(values);
  });
  if (additions.length)
    sheet
      .getRange(sheet.getLastRow() + 1, 1, additions.length, columns.length)
      .setValues(additions);
}
function findTitle_(id) {
  var t = records_("titles").find(function (r) {
    return Number(r.id) === Number(id);
  });
  if (!t) throw new Error("Title not found");
  return attachEpisodes_(t);
}
function titleId_(tmdbId, type) {
  return Number(tmdbId) * 2 + (type === "tv" ? 1 : 0);
}
function tmdb_(path) {
  var token = properties_().getProperty("TMDB_READ_TOKEN");
  if (!token) throw new Error("Set TMDB_READ_TOKEN in Script Properties");
  var r = UrlFetchApp.fetch("https://api.themoviedb.org/3" + path, {
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true,
  });
  if (r.getResponseCode() !== 200)
    throw new Error("TMDB request failed (" + r.getResponseCode() + ")");
  return JSON.parse(r.getContentText());
}
function search_(query) {
  if (typeof query !== "string" || query.length > 200)
    throw new Error("Invalid search");
  if (!query.trim()) return [];
  return tmdb_("/search/multi?query=" + encodeURIComponent(query))
    .results.filter(function (t) {
      return ["movie", "tv"].indexOf(t.media_type) >= 0;
    })
    .slice(0, 20)
    .map(function (t) {
      return {
        tmdb_id: t.id,
        media_type: t.media_type,
        title: t.title || t.name,
        year:
          Number((t.release_date || t.first_air_date || "").slice(0, 4)) ||
          null,
        poster_path: t.poster_path,
        overview: t.overview,
      };
    });
}
function hydrate_(id, type, force) {
  if (
    !Number.isSafeInteger(Number(id)) ||
    Number(id) <= 0 ||
    ["movie", "tv"].indexOf(type) < 0
  )
    throw new Error("Invalid title");
  var key = titleId_(id, type),
    old = records_("titles").find(function (t) {
      return Number(t.id) === key;
    });
  if (
    old &&
    !force &&
    Date.now() - Date.parse(old.metadata_updated_at) < 86400000
  )
    return attachEpisodes_(old);
  var d = tmdb_(
    "/" +
      type +
      "/" +
      id +
      "?append_to_response=credits,keywords,watch/providers",
  );
  var episodes = [];
  if (type === "tv") {
    var seasons = (d.seasons || []).filter(function (s) {
      return s.season_number > 0;
    });
    var requests = seasons.map(function (s) {
      return {
        url:
          "https://api.themoviedb.org/3/tv/" +
          id +
          "/season/" +
          s.season_number,
        headers: {
          Authorization:
            "Bearer " + properties_().getProperty("TMDB_READ_TOKEN"),
        },
        muteHttpExceptions: true,
      };
    });
    if (requests.length)
      UrlFetchApp.fetchAll(requests).forEach(function (r) {
        if (r.getResponseCode() !== 200)
          throw new Error("TMDB season request failed");
        var s = JSON.parse(r.getContentText());
        (s.episodes || []).forEach(function (e) {
          episodes.push({
            season_number: e.season_number,
            episode_number: e.episode_number,
            name: e.name,
            air_date: e.air_date,
          });
        });
      });
  }
  var title = {
    id: key,
    tmdb_id: Number(id),
    media_type: type,
    title: d.title || d.name,
    year:
      Number((d.release_date || d.first_air_date || "").slice(0, 4)) || null,
    runtime_min:
      d.runtime ||
      (d.episode_run_time || [])[0] ||
      (d.last_episode_to_air || {}).runtime ||
      null,
    genres: (d.genres || []).map(function (g) {
      return g.name;
    }),
    keywords: (
      (d.keywords || {}).keywords ||
      (d.keywords || {}).results ||
      []
    ).map(function (k) {
      return k.name;
    }),
    overview: d.overview,
    poster_path: d.poster_path,
    cast_top: ((d.credits || {}).cast || []).slice(0, 8).map(function (c) {
      return c.name;
    }),
    director:
      (
        ((d.credits || {}).crew || []).find(function (c) {
          return c.job === "Director";
        }) || {}
      ).name || null,
    providers: (d["watch/providers"] || {}).results || {},
    episode_data: episodes,
    metadata_updated_at: now_(),
  };
  putEpisodes_(key, episodes);
  put_("titles", key, title);
  return title;
}
function tracking_(user, titleId) {
  return records_("tracking").find(function (r) {
    return r.user_id === user.id && Number(r.title_id) === Number(titleId);
  });
}
function trackingView_(row) {
  return row
    ? Object.assign(cleanRecord_(row), { titles: findTitle_(row.title_id) })
    : null;
}
function bootstrap_(user) {
  var own = records_("tracking").filter(function (r) {
      return r.user_id === user.id;
    }),
    titles = records_("titles"),
    episodeRows = records_("episodes"),
    byId = {};
  titles.forEach(function (t) {
    byId[t.id] = Object.assign(cleanRecord_(t), {
      episode_data: episodeRows
        .filter(function (e) {
          return Number(e.title_id) === Number(t.id);
        })
        .map(function (e) {
          return {
            season_number: e.season,
            episode_number: e.episode,
            name: e.name,
            air_date: e.air_date,
          };
        }),
    });
  });
  var date = Utilities.formatDate(
      new Date(),
      user.timezone || "America/New_York",
      "yyyy-MM-dd",
    ),
    end = new Date(Date.now() + 60 * 86400000).toISOString().slice(0, 10),
    releases = [];
  own.forEach(function (row) {
    var title = byId[row.title_id];
    if (!title) return;
    var groups = {};
    (title.episode_data || []).forEach(function (e) {
      if (!e.air_date) return;
      if (!groups[e.air_date]) groups[e.air_date] = [];
      groups[e.air_date].push(e);
    });
    Object.keys(groups).forEach(function (day) {
      if (day < date || day > end) return;
      groups[day].forEach(function (e) {
        releases.push({
          title_id: title.id,
          season: e.season_number,
          episode: e.episode_number,
          episode_title: e.name,
          air_date: day,
          release_group: title.id + ":" + day,
          group_size: groups[day].length,
          titles: title,
        });
      });
    });
  });
  return {
    profile: cleanRecord_(user),
    tracking: own.map(function (r) {
      return Object.assign(cleanRecord_(r), { titles: byId[r.title_id] });
    }),
    releases: releases.sort(function (a, b) {
      return a.air_date.localeCompare(b.air_date);
    }),
    reviews: records_("import_reviews")
      .filter(function (r) {
        return r.user_id === user.id && !r.resolved_at;
      })
      .map(cleanRecord_),
    server_time: now_(),
  };
}
function commitTracking_(row) {
  row.revision = Number(row.revision || 0) + 1;
  row.updated_at = now_();
  put_("tracking", row.key, row);
  return row;
}
function reScore_(uid, tier) {
  if (!tier) return;
  var band = { loved: 0.75, liked: 0.5, fine: 0.25, nah: 0 }[tier],
    list = records_("tracking")
      .filter(function (r) {
        return r.user_id === uid && r.tier === tier;
      })
      .sort(function (a, b) {
        return a.rank_pos - b.rank_pos || a.title_id - b.title_id;
      });
  list.forEach(function (r, i) {
    r.score =
      band + 0.25 * (list.length === 1 ? 0.5 : 1 - i / (list.length - 1));
    commitTracking_(r);
  });
}
function apply_(user, title, operation, payload) {
  var key = user.id + ":" + title.id,
    row = tracking_(user, title.id),
    oldTier = row && row.tier;
  if (operation === "remove") {
    remove_("tracking", key);
    reScore_(user.id, oldTier);
    return null;
  }
  if (operation === "status") {
    if (["want", "watching", "done", "dropped"].indexOf(payload.status) < 0)
      throw new Error("Invalid status");
    row = row || {
      key: key,
      user_id: user.id,
      title_id: title.id,
      status: "want",
      last_season: 1,
      last_episode: 0,
      tier: null,
      rank_pos: null,
      score: null,
      seen_before: !!payload.seen_before,
      follow: "auto",
      revision: 0,
    };
    row.status = payload.status;
  } else {
    if (!row) throw new Error("Title not tracked");
    if (operation === "advance") {
      var date = Utilities.formatDate(
        new Date(),
        user.timezone || "America/New_York",
        "yyyy-MM-dd",
      );
      var next = (title.episode_data || [])
        .slice()
        .sort(function (a, b) {
          return (
            a.season_number - b.season_number ||
            a.episode_number - b.episode_number
          );
        })
        .find(function (e) {
          return (
            e.season_number > 0 &&
            (e.season_number > row.last_season ||
              (e.season_number === row.last_season &&
                e.episode_number > row.last_episode))
          );
        });
      if (!next || !next.air_date || next.air_date > date)
        throw new Error("No unwatched aired episode");
      row.last_season = next.season_number;
      row.last_episode = next.episode_number;
      row.status = "watching";
    } else if (operation === "position") {
      if (
        !Number.isInteger(payload.season) ||
        payload.season < 0 ||
        !Number.isInteger(payload.episode) ||
        payload.episode < 0
      )
        throw new Error("Invalid episode position");
      if (
        payload.episode > 0 &&
        !(title.episode_data || []).some(function (e) {
          return (
            e.season_number === payload.season &&
            e.episode_number === payload.episode
          );
        })
      )
        throw new Error("Episode not found in TMDB metadata");
      row.last_season = payload.season;
      row.last_episode = payload.episode;
    } else if (operation === "follow") {
      if (["auto", "always", "never"].indexOf(payload.follow) < 0)
        throw new Error("Invalid follow setting");
      row.follow = payload.follow;
    } else if (operation === "rate") {
      if (row.status !== "done") throw new Error("Mark watched before rating");
      if (["loved", "liked", "fine", "nah"].indexOf(payload.tier) < 0)
        throw new Error("Invalid tier");
      var comparisons = payload.comparisons || [];
      if (!Array.isArray(comparisons) || comparisons.length > 3)
        throw new Error("Maximum three comparisons");
      var all = records_("tracking").filter(function (r) {
        return (
          r.user_id === user.id &&
          r.tier === payload.tier &&
          r.title_id !== title.id
        );
      });
      comparisons.forEach(function (c) {
        if (
          !all.some(function (r) {
            return Number(r.title_id) === Number(c.other);
          }) ||
          [Number(title.id), Number(c.other)].indexOf(Number(c.winner)) < 0
        )
          throw new Error("Invalid comparison");
      });
      var before = all.find(function (r) {
          return Number(r.title_id) === Number(payload.before);
        }),
        after = all.find(function (r) {
          return Number(r.title_id) === Number(payload.after);
        });
      if (
        (payload.before != null && !before) ||
        (payload.after != null && !after)
      )
        throw new Error("Ranking changed. Please rate this title again");
      var lo = after ? Number(after.rank_pos) : null,
        hi = before ? Number(before.rank_pos) : null;
      if (lo !== null && hi !== null && lo >= hi)
        throw new Error("Invalid rank bounds");
      if (lo === null && hi === null)
        lo = all.reduce(function (max, r) {
          return Math.max(max, Number(r.rank_pos));
        }, 0);
      row.tier = payload.tier;
      row.rank_pos =
        lo === null ? hi - 1024 : hi === null ? lo + 1024 : (lo + hi) / 2;
    } else throw new Error("Unknown tracking operation");
  }
  commitTracking_(row);
  if (operation === "rate") {
    (payload.comparisons || []).forEach(function (c, i) {
      var cid = payload.event_key + ":" + i;
      put_("comparisons", cid, {
        id: cid,
        user_id: user.id,
        title_a: title.id,
        title_b: c.other,
        winner: c.winner,
        compared_at: now_(),
      });
    });
    reScore_(user.id, row.tier);
    if (oldTier !== row.tier) reScore_(user.id, oldTier);
    row = tracking_(user, title.id);
  }
  if (operation === "advance") {
    var eventKey = user.id + ":" + payload.event_key;
    put_("watch_events", eventKey, {
      key: eventKey,
      user_id: user.id,
      title_id: title.id,
      season: row.last_season,
      episode: row.last_episode,
      started_at: now_(),
      completed: true,
      source: "manual",
      service: "",
    });
  }
  return trackingView_(row);
}
function exactlyOnce_(user, requestId, request, fn) {
  if (
    typeof requestId !== "string" ||
    requestId.length < 8 ||
    requestId.length > 120
  )
    throw new Error("A stable request ID is required");
  var key = user.id + ":" + requestId,
    fingerprint = digest_(JSON.stringify(request)),
    previous = records_("operations").find(function (r) {
      return r.key === key;
    });
  if (previous) {
    if (previous.fingerprint !== fingerprint)
      throw new Error("Request ID was reused for a different edit");
    if (previous.status === "done")
      return previous.result && previous.result.title_id
        ? trackingView_(previous.result)
        : previous.result;
    throw new Error(
      "UNCERTAIN: this edit was interrupted. Refresh your library before making another edit.",
    );
  }
  put_("operations", key, {
    key: key,
    user_id: user.id,
    request_id: requestId,
    fingerprint: fingerprint,
    status: "started",
    result: null,
    updated_at: now_(),
  });
  SpreadsheetApp.flush(); // The started journal must persist before any mutation.
  try {
    var result = fn(),
      stored = result;
    if (result && result.title_id) {
      stored = Object.assign({}, result);
      delete stored.titles;
    }
    SpreadsheetApp.flush(); // Commit all action rows before acknowledging completion.
    put_("operations", key, {
      key: key,
      user_id: user.id,
      request_id: requestId,
      fingerprint: fingerprint,
      status: "done",
      result: stored,
      updated_at: now_(),
    });
    SpreadsheetApp.flush();
    return result;
  } catch (e) {
    // Sheets lacks transactions. A started journal entry remains on ANY failure,
    // so retry can never silently apply a second episode increment after a crash.
    throw e;
  }
}
function normalize_(s) {
  return String(s)
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[^a-z0-9 ]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}
function importRow_(user, input) {
  var raw = String(input.title || "").slice(0, 500);
  if (!raw) throw new Error("Import title is required");
  var service = String(input.service || "csv"),
    aliasKey = digest_(service + "|" + raw),
    alias = records_("aliases").find(function (r) {
      return r.key === aliasKey;
    }),
    title;
  if (alias && alias.title_id) title = findTitle_(alias.title_id);
  else {
    var query = input.query || raw,
      candidates = search_(query),
      exact = candidates.filter(function (t) {
        return (
          normalize_(t.title) === normalize_(query) &&
          (!input.year || Number(input.year) === t.year)
        );
      });
    if (exact.length === 1) {
      title = hydrate_(exact[0].tmdb_id, exact[0].media_type);
      put_("aliases", aliasKey, {
        key: aliasKey,
        raw_string: raw,
        service: service,
        title_id: title.id,
        confidence: 1,
        confirmed_by_user: false,
      });
    } else {
      var reviewKey =
          user.id +
          ":" +
          digest_(service + "|" + raw + "|" + (input.watched_at || "")),
        existing = records_("import_reviews").find(function (r) {
          return r.key === reviewKey;
        });
      if (!existing)
        put_("import_reviews", reviewKey, {
          id: reviewKey,
          key: reviewKey,
          user_id: user.id,
          raw_string: raw,
          service: service,
          watched_at: input.watched_at,
          season: input.season,
          episode: input.episode,
          candidates: candidates,
          resolved_at: null,
        });
      return { review: true };
    }
  }
  if (!tracking_(user, title.id))
    apply_(user, title, "status", { status: "done", seen_before: true });
  var eventKey =
    user.id +
    ":" +
    digest_(service + "|" + raw + "|" + (input.watched_at || ""));
  if (
    !records_("watch_events").some(function (r) {
      return r.key === eventKey;
    })
  )
    put_("watch_events", eventKey, {
      key: eventKey,
      user_id: user.id,
      title_id: title.id,
      season: input.season,
      episode: input.episode,
      started_at: input.watched_at || now_(),
      completed: true,
      source: "csv",
      service: service,
    });
  return { imported: true };
}
function dispatch_(user, b) {
  switch (b.action) {
    case "bootstrap":
      return bootstrap_(user);
    case "search":
      return search_(b.query);
    case "detail":
      return hydrate_(b.tmdb_id, b.media_type);
    case "track":
      return exactlyOnce_(user, (b.payload || {}).event_key, b, function () {
        var title = b.title_id
          ? findTitle_(b.title_id)
          : hydrate_(b.tmdb_id, b.media_type);
        return apply_(user, title, b.operation, b.payload || {});
      });
    case "profile":
      return exactlyOnce_(user, b.request_id, b, function () {
        var p = b.profile || {};
        if (p.display_name !== undefined) {
          if (
            typeof p.display_name !== "string" ||
            !p.display_name.trim() ||
            p.display_name.length > 80
          )
            throw new Error("Invalid display name");
          user.display_name = p.display_name.trim();
        }
        if (p.region !== undefined) {
          if (!/^[A-Z]{2}$/.test(p.region)) throw new Error("Invalid region");
          user.region = p.region;
        }
        if (p.timezone !== undefined) {
          Utilities.formatDate(new Date(), p.timezone, "yyyy-MM-dd");
          user.timezone = p.timezone;
        }
        if (p.streaming_services !== undefined) {
          if (
            !Array.isArray(p.streaming_services) ||
            p.streaming_services.length > 12 ||
            p.streaming_services.some(function (s) {
              return (
                ["netflix", "max", "prime", "apple", "hulu", "disney"].indexOf(
                  s,
                ) < 0
              );
            })
          )
            throw new Error("Invalid services");
          user.streaming_services = p.streaming_services;
        }
        put_("users", user.id, user);
        return cleanRecord_(user);
      });
    case "import":
      if (!Array.isArray(b.rows) || b.rows.length > 5)
        throw new Error("Import at most five rows per batch");
      return b.rows.map(function (r) {
        return importRow_(user, r);
      });
    case "review":
      var review = records_("import_reviews").find(function (r) {
        return r.id === b.id && r.user_id === user.id;
      });
      if (!review) throw new Error("Review not found");
      var title = hydrate_(b.tmdb_id, b.media_type),
        aliasKey = digest_(review.service + "|" + review.raw_string);
      put_("aliases", aliasKey, {
        key: aliasKey,
        raw_string: review.raw_string,
        service: review.service,
        title_id: title.id,
        confidence: 1,
        confirmed_by_user: true,
      });
      importRow_(user, {
        title: review.raw_string,
        service: review.service,
        watched_at: review.watched_at,
        season: review.season,
        episode: review.episode,
      });
      review.resolved_at = now_();
      put_("import_reviews", review.id, review);
      return { ok: true };
    case "refresh-metadata":
      var own = records_("tracking").filter(function (r) {
          return r.user_id === user.id;
        }),
        ids = b.title_ids || [];
      if (!Array.isArray(ids) || ids.length > 3)
        throw new Error("Refresh at most three titles per request");
      ids.forEach(function (id) {
        if (
          !own.some(function (r) {
            return Number(r.title_id) === Number(id);
          })
        )
          throw new Error("Title not in your library");
        var t = findTitle_(id);
        hydrate_(t.tmdb_id, t.media_type, true);
      });
      return { ok: true };
    default:
      throw new Error("This feature is deferred in the Sheets version");
  }
}
