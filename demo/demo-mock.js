/* =========================================================================
   DEMO MODE - replaces Microsoft sign-in (MSAL) and Microsoft Graph/
   SharePoint with an in-memory fake, so the real admin app code runs
   unmodified against INVENTED sample data. Nothing leaves the browser,
   nothing is saved (a reload resets everything), and no real person's
   information is in this file - every name, number and note below is made
   up. Loaded in place of the MSAL <script> tag (see build-demo.ps1).
========================================================================= */
(function(){
  // ---- tiny seeded RNG so the sample data is the same on every load -----
  let seed = 20260930;
  const rnd = () => { seed = (seed * 1664525 + 1013904223) % 4294967296; return seed / 4294967296; };
  const pick = a => a[Math.floor(rnd()*a.length)];
  const pad = n => String(n).padStart(2,'0');
  const iso = d => `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
  const dt = (d, hm) => `${iso(d)}T${hm}:00`;
  const addDays = (d, n) => { const x = new Date(d); x.setDate(x.getDate()+n); return x; };

  const today = new Date(); today.setHours(0,0,0,0);
  const year = today.getFullYear();

  // ---- sample people (all invented) --------------------------------------
  const ministerSeed = [
    ['Grace Hamilton','Female'],['Daniel Okafor','Male'],['Miriam Chen','Female'],['Samuel Rivera','Male'],
    ['Ruth Anderson','Female'],['Joel Thompson','Male'],['Naomi Patel','Female'],['Caleb Brooks','Male'],
    ['Esther Kim','Female'],['Nathan Wright','Male'],['Priya Menon','Female'],['Marcus Lee','Male']
  ];
  const recipients = ['Alex Morgan','Jordan Blake','Taylor Quinn','Casey Rowe','Riley Stone','Morgan Ellis','Jamie Frost','Avery Cole','Parker Hale','Quinn Sutton','Rowan Pike','Sage Whitlock','Skylar Dunn','Reese Calloway','Emerson Vale','Harper Lowell','Finley Ash','Dakota Reyes','Logan Marsh','Peyton Hart','Cameron Boyd','Drew Sinclair','Blair Winters','Kendall Fox','Lane Holloway','Shea Barlow','Tatum Reid','Wren Castillo','Remy Delacroix','Arden Price','Sloane Mercer','Ellis Navarro','Marlow Grant','Sutton Beck','Hollis Crane','Lennox Pratt','Oakley Burke','Bellamy Shaw','Jules Archer','Nico Fairbanks','Kai Rutherford','Ari Whitaker','Sasha Lindgren','Devon Maddox','Corey Linden','Briar Kessler','Tristan Voss','Mara Ibsen','Zane Holt','Lark Donovan'];
  const locationNames = ['Chapel Room','Prayer Room A','Prayer Room B','Online (Zoom)'];

  const ministers = ministerSeed.map(([n,g],i)=>({
    id: 100+i,
    fields: { Title:n, Email:`${n.toLowerCase().replace(/[^a-z]+/g,'.')}@example.com`, Phone:`555-01${pad(i)}`, SignInEmail:`${n.toLowerCase().replace(/[^a-z]+/g,'.')}@demo.example`, Status: i===11?'Inactive':'Active', Gender:g }
  }));
  const mById = Object.fromEntries(ministers.map(m=>[m.id,m]));

  // ---- sessions -----------------------------------------------------------
  const sessions = [];
  let sid = 1000;
  const twoMinisters = (gender) => {
    const pool = ministers.filter(m=>m.fields.Status==='Active' && (!gender || m.fields.Gender===gender));
    const a = pick(pool); let b = pick(pool); while(b.id===a.id) b = pick(pool);
    return [a.id, b.id];
  };
  const addSession = (o) => {
    const id = sid++;
    const d = o.date;
    sessions.push({ id, createdDateTime: dt(addDays(d,-9),'09:00'), fields: Object.assign({
      Title: `${o.name} - ${iso(d)}`, RecipientName:o.name, RecipientContact:o.contact||'', RecipientGender:o.gender||'', GenderPreference:'Mixed',
      SessionDate: dt(d,o.start||'13:00'), SessionEndDate: dt(d,o.end||'14:30'),
      AssignedMinisterIDs:(o.team||[]).join(','), LeadMinisterIDs:(o.team||[])[0]||'',
      LocationName:o.loc||pick(locationNames), Notes:o.notes||'', ApptType:o.type||'First', Status:o.status||'Completed',
      PreviousSessionId:o.prev||null, WaCreated:!!o.wa, WaLink:o.wa?'https://example.com/chat-demo':'', WaNotified:!!o.wa,
      Priority:false, Contacted:false, ConfirmedMinisterIDs:o.confirmed||'', DeclinedMinisterIDs:o.declined||''
    }, o.extra||{}) });
    return id;
  };

  // Completed Freedom Sessions across the year so far (first + follow-ups)
  const tuesdaysBefore = [];
  for(let d = new Date(year,0,6); d < addDays(today,-3); d = addDays(d,7)) if(d.getDay()===2) tuesdaysBefore.push(new Date(d));
  const gendersOf = n => (n.charCodeAt(0)%2 ? 'Female':'Male');
  recipients.slice(0,36).forEach((name,i)=>{
    const g = gendersOf(name);
    const first = tuesdaysBefore[Math.floor(rnd()*tuesdaysBefore.length)];
    let prev = null; let d = first;
    const n = 1 + Math.floor(rnd()*3);
    for(let k=0;k<n && d < addDays(today,-3);k++){
      prev = addSession({ name, gender:g, contact:`${name.split(' ')[0].toLowerCase()}@example.com`, date:d, team:twoMinisters(), type:k===0?'First':'Follow-up', prev, wa:rnd()>0.4, notes:k===0?'Requested prayer for a season of change.':'Follow-up conversation.' });
      d = addDays(d, 7*(2+Math.floor(rnd()*3)));
    }
  });
  // Sunday Drop-In (4th Sunday) + Training, completed
  for(let m=0;m<today.getMonth()+1;m++){
    const d = new Date(year,m,1); while(d.getDay()!==0) d.setDate(d.getDate()+1); d.setDate(d.getDate()+21);
    if(d < today) addSession({ name:'Sunday Drop-In', date:d, team:twoMinisters(), type:'Drop-in', start:'17:00', end:'19:00', loc:'Chapel Room' });
    const t = addDays(new Date(year,m,1), 9+Math.floor(rnd()*10));
    if(t < today && m%2===0) addSession({ name:'Training', date:t, team:twoMinisters(), type:'Training', start:'18:30', end:'20:00', loc:'Prayer Room A' });
  }

  // Upcoming scheduled sessions (some confirmed / declined by ministers)
  const upcoming = [];
  [3,5,8,10].forEach((off,i)=>{
    let d = addDays(today, off); while(d.getDay()!==2) d = addDays(d,1);
    const team = twoMinisters();
    const name = recipients[36+i];
    const id = addSession({ name, gender:gendersOf(name), contact:`${name.split(' ')[0].toLowerCase()}@example.com`, date:d, team, status:'Scheduled', wa:true,
      notes:'First session booked from the waiting list.', confirmed: i===0?String(team[0]):(i===1?team.join(','):''), declined: i===2?String(team[1]):'' });
    upcoming.push({ id, date:d, team });
  });

  // Waiting list (some priority / contacted)
  const waiting = [];
  [40,41,42,43,44].forEach((ri,i)=>{
    const name = recipients[ri]; const d = addDays(today, -4 - i*5);
    const id = addSession({ name, gender:gendersOf(name), contact:`${name.split(' ')[0].toLowerCase()}@example.com`, date:d, status:'Waiting', type:'First', notes:i===1?'Prefers an evening time.':'', extra:{ SessionDate:null, SessionEndDate:null, LocationName:'', AssignedMinisterIDs:'', LeadMinisterIDs:'', Priority:i===0, Contacted:i===2 } });
    sessions[sessions.length-1].createdDateTime = dt(d,'10:00');
    waiting.push(id);
  });

  // ---- planning slots (next three Tuesdays) -------------------------------
  const slots = []; let slid = 3000;
  for(let w=0; w<3; w++){
    let d = addDays(today, 1 + w*7); while(d.getDay()!==2) d = addDays(d,1);
    for(let k=0;k<3;k++) slots.push({ id: slid++, fields:{ SlotDate: dt(d,'00:00'), StartTime: k===0?'13:00':(k===1?'14:00':'18:00'), EndTime: k===0?'14:00':(k===1?'15:00':'19:00'), Status:'Open', ClaimedWaitingId:null, ClaimedDate:null, ProposedMinisterIDs:'', ProposedLeadIDs:'', LinkedSessionId:null, Notes:'' } });
  }
  // one tentative claim and one booked slot, so every state shows up
  Object.assign(slots[1].fields, { Status:'Tentative', ClaimedWaitingId:waiting[1], ClaimedDate:dt(today,'09:00'), ProposedMinisterIDs:'100,101', ProposedLeadIDs:'100' });
  Object.assign(slots[3].fields, { Status:'Booked', LinkedSessionId:upcoming[0].id, ProposedMinisterIDs:upcoming[0].team.join(','), ProposedLeadIDs:String(upcoming[0].team[0]), SlotDate: dt(upcoming[0].date,'00:00') });

  // ---- blackout dates ----------------------------------------------------
  const blackouts = [
    { id:4000, createdDateTime: dt(addDays(today,-10),'08:00'), createdBy:{user:{email:ministers[0].fields.SignInEmail}}, fields:{ Title:ministers[0].fields.Title, BlackoutDate: dt(addDays(today,6),'00:00'), EndDate: dt(addDays(today,13),'00:00'), Notes:'Family trip' } },
    { id:4001, createdDateTime: dt(addDays(today,-8),'08:00'), createdBy:{user:{email:ministers[3].fields.SignInEmail}}, fields:{ Title:ministers[3].fields.Title, BlackoutDate: dt(addDays(today,9),'00:00'), EndDate:null, Notes:'Work conference' } },
    { id:4002, createdDateTime: dt(addDays(today,-3),'08:00'), createdBy:{user:{email:'admin@demo.example'}}, fields:{ Title:ministers[5].fields.Title, BlackoutDate: dt(addDays(today,16),'00:00'), EndDate:null, Notes:'Church closed (holiday)' } }
  ];

  // ---- intake forms (invented answers) -----------------------------------
  const intake = []; let iid = 5000;
  const mkIntake = (name, status, daysAgo) => {
    const id = iid++;
    const email = `${name.split(' ')[0].toLowerCase()}@example.com`;
    intake.push({ id, createdDateTime: dt(addDays(today,-daysAgo),'11:00'), fields:{ Title:name, Token:`demo-${id}`, RecipientName:name, RecipientEmail:email, Status:status,
      SubmittedAt: status==='InProgress' ? null : dt(addDays(today,-daysAgo),'11:30'),
      ResponsesJSON: JSON.stringify({ name, email, age:'30-39' }), SignatureDataUrl:'' } });
    return id;
  };
  const linkA = mkIntake(sessions.find(s=>s.fields.Status==='Scheduled').fields.RecipientName,'Submitted',6);
  sessions.find(s=>s.fields.Status==='Scheduled').fields.IntakeResponseId = linkA;
  mkIntake(recipients[41],'Submitted',5); mkIntake(recipients[42],'Submitted',3); mkIntake(recipients[43],'Submitted',2);
  mkIntake(recipients[44],'InProgress',1); mkIntake(recipients[45],'InProgress',2); mkIntake(recipients[46],'Submitted',60);
  intake[intake.length-1].fields.Status = 'Archived';

  const lists = {
    PrayerMinisters: ministers,
    PrayerSessions: sessions,
    BlackoutDates: blackouts,
    Locations: locationNames.map((n,i)=>({ id:2000+i, fields:{ Title:n } })),
    PlanningSlots: slots,
    TimeWindows: [
      { id:6000, fields:{ Title:'Tuesday', Kind:'Weekday', DayOfWeek:2, StartTime:'13:00', EndTime:'15:00', LocationName:'' } },
      { id:6001, fields:{ Title:'Drop-In', Kind:'DropIn', DayOfWeek:null, StartTime:'17:00', EndTime:'19:00', LocationName:'Chapel Room' } },
      { id:6002, fields:{ Title:'Default Location', Kind:'DefaultLocation', DayOfWeek:null, StartTime:'', EndTime:'', LocationName:'Chapel Room' } }
    ],
    IntakeResponses: intake,
    IntakeFormSchema: [ { id:7000, fields:{ Title:'schema', SchemaVersion:1, get SchemaJSON(){ return JSON.stringify(window.FPHM_INTAKE_CONFIG); } } } ]
  };
  const listIdByName = Object.fromEntries(Object.keys(lists).map((n,i)=>[n,'list-'+n]));
  let nextId = 9000;

  // ---- fake fetch: Graph -> in-memory lists; other hosts -> harmless ok ---
  const realFetch = window.fetch.bind(window);
  const json = (obj, status=200) => Promise.resolve(new Response(status===204?null:JSON.stringify(obj), { status, headers:{'Content-Type':'application/json'} }));
  window.fetch = function(url, opts){
    const u = String(url);
    const method = ((opts && opts.method) || 'GET').toUpperCase();
    if(u.startsWith('https://graph.microsoft.com/v1.0')){
      const path = decodeURIComponent(u.slice('https://graph.microsoft.com/v1.0'.length));
      let m;
      if((m = path.match(/^\/sites\/[^/]+:\/.+$/))) return json({ id:'demo-site' });
      if(/^\/sites\/demo-site\/lists\?/.test(path)) return json({ value: Object.keys(lists).map(n=>({ id:listIdByName[n], displayName:n })) });
      if((m = path.match(/^\/sites\/demo-site\/lists\/list-(\w+)\/items\?/)) && method==='GET'){
        return json({ value: lists[m[1]].map(it=>({ id:String(it.id), createdDateTime: it.createdDateTime||'2026-01-02T09:00:00Z', createdBy: it.createdBy||{user:{email:'admin@demo.example'}}, fields: it.fields })) });
      }
      if((m = path.match(/^\/sites\/demo-site\/lists\/list-(\w+)\/items$/)) && method==='POST'){
        const body = JSON.parse(opts.body); const id = nextId++;
        lists[m[1]].push({ id, createdDateTime:new Date().toISOString(), fields: body.fields });
        return json({ id:String(id) }, 201);
      }
      if((m = path.match(/^\/sites\/demo-site\/lists\/list-(\w+)\/items\/(\d+)\/fields$/)) && method==='PATCH'){
        const it = lists[m[1]].find(x=>String(x.id)===m[2]); if(it) Object.assign(it.fields, JSON.parse(opts.body));
        return json(it ? it.fields : {});
      }
      if((m = path.match(/^\/sites\/demo-site\/lists\/list-(\w+)\/items\/(\d+)$/)) && method==='DELETE'){
        const arr = lists[m[1]]; const ix = arr.findIndex(x=>String(x.id)===m[2]); if(ix>=0) arr.splice(ix,1);
        return json(null, 204);
      }
      // account creation / group membership etc: pretend it worked
      return method==='DELETE' || /\/\$ref$/.test(path) ? json(null,204) : json({ id:'demo-user' });
    }
    if(/^https?:\/\//.test(u) && !u.startsWith(location.origin) && !/fonts\.(googleapis|gstatic)\.com|cdn\./.test(u)) return json({ ok:true });
    return realFetch(url, opts);
  };

  // ---- fake MSAL: always "signed in" as an invented admin -----------------
  const demoAccount = { homeAccountId:'demo', environment:'demo', tenantId:'demo', username:'admin@demo.example', localAccountId:'demo', name:'Demo Admin' };
  let active = demoAccount, signedIn = true;
  window.msal = { PublicClientApplication: function(){
    return {
      initialize: () => Promise.resolve(),
      getAllAccounts: () => signedIn ? [demoAccount] : [],
      getActiveAccount: () => signedIn ? active : null,
      setActiveAccount: a => { active = a; },
      acquireTokenSilent: () => Promise.resolve({ accessToken:'demo-token', account:demoAccount }),
      acquireTokenPopup: () => Promise.resolve({ accessToken:'demo-token', account:demoAccount }),
      loginPopup: () => { signedIn = true; return Promise.resolve({ account:demoAccount, accessToken:'demo-token' }); },
      logoutPopup: () => { signedIn = false; return Promise.resolve(); },
      clearCache: () => Promise.resolve()
    };
  } };
  // Switch User reloads the page, which resets the demo - nothing to keep.
})();
