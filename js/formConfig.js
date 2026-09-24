/* =========================================================================
   FORM CONFIG — the intake questionnaire's questions and flow.

   This is the ONE place to edit when the ministry wants to add, remove,
   reorder or rephrase a question. Both the public intake form (index.html)
   and the staff view (admin/index.html) render from this file, so a change
   here shows up in both places automatically.

   Shape:
     SECTIONS: array of { id, title, intro?, visibleIf?, questions: [...] }
     - A section with `visibleIf` is skipped entirely (like a page skip in
       the original Microsoft Form) when the condition doesn't match.
     - Each question: { id, type, label, options?, otherOption?, required?,
       visibleIf?, min?, max?, placeholder? }
     - `id` must be unique across the WHOLE form (answers are stored as a
       flat object keyed by id).
     - `visibleIf` shapes (evaluated against the current answers object):
         { id, equals: value }
         { id, in: [values] }
         { id, includes: value }   -- for checkbox-array answers
         { id, notEquals: value }
     - `type`: 'text' | 'email' | 'number' | 'textarea' | 'radio' |
                'select' | 'checkboxes'
     - `otherOption: true` on a checkboxes question adds a final "Other"
       choice with its own free-text box.
========================================================================= */

(function () {

const INTRO_TEXT = {
  title: "Freedom Prayer Healing Ministry Intake Questionnaire",
  body: [
    `When you submit this form, it will not automatically collect your details like name and email address unless you provide it yourself.`,
    `As you read through this questionnaire, you may find yourself thinking, "Why so many questions?" or "Why such personal ones?". After all, you may be seeking prayer for a specific issue - whether physical healing, emotional healing, or oppression by evil spirits - and yet the questions cover many areas.`,
    `To help you understand the reason for this, here is our framework for praying for healing.`,
    `The world and life are not just physical. Spiritual forces are real.`,
    `God gives life, and His nature and desire is to heal, rescue and redeem. Although He allows sickness and oppression, He doesn't generally cause them. This means that we can confidently pray for healing and freedom. Sickness — whether physical, emotional, mental, or spiritual — does not come from God. It comes from the Fall, spiritual attack, or the consequences of our choices.`,
    `The enemy want to kill, steal and destroy. Unclean spirits are like rats drawn to garbage — the "garbage" being areas of sin, trauma, or brokenness. These entry points or open doors may come from personal choices, things done to us, or generational patterns. The enemy uses any access he can get to oppress and harass.`,
    `People are made up of body, soul (mind, emotions, will), and spirit — and all of these are connected. What affects one part can impact the others. For example, emotional stress can lead to physical illness, and physical illness can affect mental and emotional health. When we pray for healing and deliverance we focus on both the symptoms (fruit) and the deeper causes (roots). Sometimes these root causes have provided "open doors" or ways that the enemy has been able to gain access to us to oppress and harass.`,
    `With this framework in mind, the questionnaire serves two purposes: To help you reflect on your life, choices, and family background. To help the prayer team identify possible root causes or open doors, so we can ask the Holy Spirit to show what needs to be addressed.`,
    `Instructions: Take your time in filling out this form. If you don't get it done all at once, you can save it and come back to it later using the "Save & Continue Later" button — you'll get a link to return to your answers.`,
    `Your thoughtful and honest answers provide a foundation for healing. Some of the questions are quite personal but please know that purpose is not to blame or condemn. Remember that the enemy loves to use shame to isolate us and keep us from living in the freedom that Jesus promises. It's when we bring our secrets in to the light that Jesus can heal our hurts and memories and forgive our sin. He really wants you to walk in freedom! What you write remains confidential to the team praying with you and to the pastor who oversees Freedom Prayer Healing Ministry.`
  ]
};

const LIABILITY_TEXT = {
  title: "Liability Release",
  body: `By signing below you indicate your agreement with this liability release.

I, the undersigned, release Creekside Church and the volunteers of Freedom Prayer Healing Ministry from any liability for any harm or perceived harm resulting from my voluntary receiving of free prayer on this and any subsequent visits. I understand that Freedom Prayer Healing Ministry is staffed by volunteers. They are not trained or licensed professionals in counseling, therapy or medical services and they are not offering medical, health care, legal or financial advice. I understand that any results of prayer received should be confirmed by my healthcare provider before changes to medical treatment or therapy that I may be receiving are made. I understand that this form and all the data on it is the sole property of Freedom Prayer Healing Ministry (Creekside Church) and that all content will be held in confidence for the purpose of praying for me.`
};

const YES_NO = ["Yes", "No"];

const SECTIONS = [
  // ---------------------------------------------------------------- 1
  {
    id: "personal",
    title: "Personal Information",
    questions: [
      { id: "name", type: "text", label: "Name", required: true },
      { id: "age", type: "radio", label: "Age", required: true,
        options: ["20-29", "30-39", "40-49", "50-59", "60-69", "70+"] },
      { id: "sex", type: "radio", label: "Sex", required: true, options: ["Male", "Female"] },
      { id: "occupation", type: "text", label: "Occupation", required: true },
      { id: "countryOfBirth", type: "text", label: "What is your country of birth?", required: true },
      { id: "email", type: "email", label: "Email", required: true }
    ]
  },
  // ---------------------------------------------------------------- 2
  {
    id: "prayerRequests",
    title: "Prayer requests",
    questions: [
      { id: "prayerFor", type: "textarea", label: "What would you like prayer for?", required: true },
      { id: "priorPrayerMinistry", type: "radio", label: "Have you ever had prayer ministry for these issues?", required: true, options: YES_NO },
      { id: "underMedicalCare", type: "radio", label: "Are you currently under medical care for these issues?", required: true, options: YES_NO },
      { id: "inCounselling", type: "radio", label: "Are you currently in counselling or psychotherapy for these issues?", required: true, options: YES_NO },
      { id: "anythingElse", type: "textarea", label: "Anything you want us to know", required: false }
    ]
  },
  // ---------------------------------------------------------------- 3
  {
    id: "family",
    title: "Family Information",
    questions: [
      { id: "parentsWantedYou", type: "radio", label: "Did your biological parents want you?", required: true, options: YES_NO },
      { id: "adopted", type: "radio", label: "Were you adopted?", required: true, options: YES_NO },
      { id: "primaryCaregiver", type: "text", label: "Who was your primary caregiver?", required: true },
      { id: "homeSecurity", type: "radio", label: "Was there a sense of security and harmony in your home during the first 12 years of your life?", required: true, options: YES_NO },
      { id: "hasSiblings", type: "radio", label: "Do you have siblings?", required: true, options: YES_NO },
      { id: "siblingsInfo", type: "textarea", label: "What would you like us to know about your siblings? Conflict?", required: true,
        visibleIf: { id: "hasSiblings", equals: "Yes" } }
    ]
  },
  // ---------------------------------------------------------------- 4
  {
    id: "parents",
    title: "Parents",
    questions: [
      { id: "motherLiving", type: "radio", label: "Is your mother still living?", required: true, options: ["Yes", "No", "Don't know"] },
      { id: "motherRelationship", type: "textarea", label: "Describe your relationship with your mother", required: true },
      { id: "fatherLiving", type: "radio", label: "Is your father still living?", required: true, options: ["Yes", "No", "Don't know"] },
      { id: "fatherRelationship", type: "textarea", label: "Describe your relationship with your father", required: true },
      { id: "parentsModeledJesus", type: "textarea", label: "Did your parents follow and model the teachings of Jesus? Please describe.", required: true },
      { id: "otherParentalFigures", type: "radio", label: "Are there any other significant parental figures in your life?", required: true, options: YES_NO },
      { id: "otherParentalFiguresExplain", type: "textarea", label: "Please explain", required: true,
        visibleIf: { id: "otherParentalFigures", equals: "Yes" } }
    ]
  },
  // ---------------------------------------------------------------- 5: marital status branch
  {
    id: "marriageStatus",
    title: "Marriage Information",
    questions: [
      { id: "maritalStatus", type: "select", label: "Marital status", required: true,
        options: ["Married", "Divorced", "Single", "Common-law", "Other"] },
      { id: "maritalStatusOther", type: "text", label: "Please specify", required: true,
        visibleIf: { id: "maritalStatus", equals: "Other" } }
    ]
  },
  {
    id: "marriageMarried",
    title: "Marriage Information (Married)",
    visibleIf: { id: "maritalStatus", equals: "Married" },
    questions: [
      { id: "m_spouseName", type: "text", label: "Name of spouse", required: false },
      { id: "m_spouseAge", type: "number", label: "Spouse's Age", required: false, min: 0, max: 120 },
      { id: "m_spouseAgeAtMarriage", type: "number", label: "How old was your spouse when you married?", required: true, min: 0, max: 120 },
      { id: "m_ageAtMarriage", type: "number", label: "How old were you when you married?", required: true, min: 0, max: 120 },
      { id: "m_separatedOrDivorced", type: "radio", label: "Have you ever been separated or divorced?", required: true, options: ["Yes", "No", "In progress"] },
      { id: "m_whoInitiated", type: "radio", label: "Who initiated it?", required: true, options: ["Me", "Spouse"],
        visibleIf: { id: "m_separatedOrDivorced", in: ["Yes", "In progress"] } },
      { id: "m_initiatedExplain", type: "textarea", label: "Please explain", required: true,
        visibleIf: { id: "m_separatedOrDivorced", in: ["Yes", "In progress"] } },
      { id: "m_previousMarriages", type: "textarea", label: "Provide brief information about any previous marriages (how long it/they lasted; reason it/they ended, conflict etc.)", required: false }
    ]
  },
  {
    id: "marriageCommonLaw",
    title: "Marriage Information (Common-law)",
    visibleIf: { id: "maritalStatus", equals: "Common-law" },
    questions: [
      { id: "cl_partnerName", type: "text", label: "Name of partner", required: false },
      { id: "cl_partnerAge", type: "number", label: "Partner's Age", required: false, min: 0, max: 120 },
      { id: "cl_partnerAgeAtMoveIn", type: "number", label: "How old was your partner when you moved in together?", required: true, min: 0, max: 120 },
      { id: "cl_ageAtMoveIn", type: "number", label: "How old were you when you moved in together?", required: true, min: 0, max: 120 },
      { id: "cl_separatedOrDivorced", type: "radio", label: "Have you ever been separated or divorced?", required: true, options: ["Yes", "No", "In progress"] },
      { id: "cl_whoInitiated", type: "radio", label: "Who initiated it?", required: true, options: ["Me", "Spouse"],
        visibleIf: { id: "cl_separatedOrDivorced", in: ["Yes", "In progress"] } },
      { id: "cl_initiatedExplain", type: "textarea", label: "Please explain", required: true,
        visibleIf: { id: "cl_separatedOrDivorced", in: ["Yes", "In progress"] } },
      { id: "cl_previousMarriages", type: "textarea", label: "Provide brief information about any previous marriages (how long it/they lasted; reason it/they ended, conflict etc.)", required: false }
    ]
  },
  {
    id: "marriageSingle",
    title: "Marriage Information (Single)",
    visibleIf: { id: "maritalStatus", equals: "Single" },
    questions: [
      { id: "s_everMarried", type: "radio", label: "Have you ever been married?", required: true, options: YES_NO },
      { id: "s_previousMarriages", type: "textarea", label: "Provide brief information about any previous marriages (how long it/they lasted; reason it/they ended, conflict etc.)", required: true,
        visibleIf: { id: "s_everMarried", equals: "Yes" } }
    ]
  },
  {
    id: "marriageDivorced",
    title: "Marriage Information (Divorced)",
    visibleIf: { id: "maritalStatus", equals: "Divorced" },
    questions: [
      { id: "d_spouseName", type: "text", label: "Name of spouse", required: false },
      { id: "d_spouseAge", type: "number", label: "Spouse's Age", required: false, min: 0, max: 120 },
      { id: "d_spouseAgeAtMarriage", type: "number", label: "How old was your spouse when you married?", required: true, min: 0, max: 120 },
      { id: "d_ageAtMarriage", type: "number", label: "How old were you when you married?", required: true, min: 0, max: 120 },
      { id: "d_whoInitiated", type: "radio", label: "Who initiated it?", required: true, options: ["Me", "Spouse"] },
      { id: "d_initiatedExplain", type: "textarea", label: "Please explain", required: true },
      { id: "d_previouslyMarried", type: "radio", label: "Have you been previously married?", required: true, options: YES_NO },
      { id: "d_previousMarriages", type: "textarea", label: "Provide brief information about any previous marriages (how long it/they lasted; reason it/they ended, conflict etc.)", required: true,
        visibleIf: { id: "d_previouslyMarried", equals: "Yes" } }
    ]
  },
  // ---------------------------------------------------------------- 6
  {
    id: "children",
    title: "Child Information",
    questions: [
      { id: "hasChildren", type: "radio", label: "Do you have or have you had any children?", required: true, options: YES_NO },
      { id: "childrenInfo", type: "textarea", label: "Please provide Name(s), Age(s), Sex(es) and if they are living.", required: true,
        visibleIf: { id: "hasChildren", equals: "Yes" } },
      { id: "lostPregnancy", type: "radio", label: "Have you ever lost a pregnancy? (miscarriage, still-born)", required: true, options: YES_NO },
      { id: "lostPregnancyExplain", type: "textarea", label: "Please explain", required: true,
        visibleIf: { id: "lostPregnancy", equals: "Yes" } },
      { id: "abortion", type: "radio", label: "Have you ever had an abortion?", required: true, options: YES_NO }
    ]
  },
  // ---------------------------------------------------------------- 7
  {
    id: "health",
    title: "Health Information",
    questions: [
      { id: "physicalHealth", type: "checkboxes", label: "Physical health - problems you experience or have experienced. Select all that apply.", required: true, otherOption: true,
        options: ["Cancer", "Skin disease", "Heart disease", "Diabetes", "Fibromyalgia", "Chronic Fatigue", "Food allergies", "None"] },
      { id: "mentalHealth", type: "checkboxes", label: "Mental health - problems you experience or have experienced. Select all that apply.", required: true, otherOption: true,
        options: ["Depression", "Anxiety", "Panic attacks", "Obsessive Compulsive Disorder (OCD)", "Eating disorder", "None"] },
      { id: "suicideHistory", type: "checkboxes", label: "Have you experienced any of the following?", required: true,
        options: ["Suicide ideation", "Suicide attempt", "None"] },
      { id: "familyHealthPatterns", type: "radio", label: "Do you see any patterns of physical or mental health problems in your parents, grandparents, and great-grandparents?", required: true, options: YES_NO },
      { id: "familyHealthPatternsExplain", type: "textarea", label: "Please explain", required: true,
        visibleIf: { id: "familyHealthPatterns", equals: "Yes" } }
    ]
  },
  // ---------------------------------------------------------------- 8
  {
    id: "christianExperience",
    title: "Christian experience",
    questions: [
      { id: "committedToJesus", type: "radio", label: "Have you ever made a commitment to follow Jesus?", required: true, options: YES_NO },
      { id: "relationshipWithJesus", type: "radio", label: "What is the current state of your relationship with Jesus?", required: true, options: ["Distant", "Luke-warm", "Close"] },
      { id: "emotionallyHonestWithGod", type: "radio", label: "Are you emotionally honest with God?", required: true, options: YES_NO },
      { id: "prayRegularly", type: "radio", label: "Do you pray regularly?", required: true, options: YES_NO },
      { id: "prayingDifficult", type: "radio", label: "Do you find praying difficult?", required: true, options: YES_NO },
      { id: "prayingDifficultExplain", type: "textarea", label: "Please explain", required: true,
        visibleIf: { id: "prayingDifficult", equals: "Yes" } },
      { id: "readBibleFrequency", type: "radio", label: "How often do you read the Bible? (Remember, this is not about feeling guilty.)", required: true, options: ["Daily", "Weekly", "Monthly", "Never"] },
      { id: "attendsChurch", type: "radio", label: "Do you presently attend a church?", required: true, options: YES_NO },
      { id: "whatChurch", type: "text", label: "What church?", required: true,
        visibleIf: { id: "attendsChurch", equals: "Yes" } },
      { id: "childhoodChurch", type: "radio", label: "Did you attend church in childhood?", required: true, options: YES_NO },
      { id: "childhoodChurchDenomination", type: "text", label: "What church (denomination)?", required: true,
        visibleIf: { id: "childhoodChurch", equals: "Yes" } },
      { id: "baptized", type: "radio", label: "Have been baptized?", required: true, options: YES_NO },
      { id: "recentChristianChanges", type: "textarea", label: "Explain recent changes in your Christian experience, if any.", required: true }
    ]
  },
  // ---------------------------------------------------------------- 9
  {
    id: "otherSpiritual",
    title: "Other religious/spiritual experiences",
    questions: [
      { id: "nonChristianReligions", type: "checkboxes", label: "Non-Christian religions - Have you been involved in any of the following?", required: true,
        options: ["Bahá'í", "Buddhism", "Hinduism", "Islam", "Jainism", "Santeria", "Satanism", "Voodoo", "Wicca", "None of the above"] },
      { id: "christianCults", type: "checkboxes", label: "Christian cults - Have you been involved in any of the following? Select all that apply.", required: true,
        options: ["Christian Science", "Jehovah's Witnesses", "Mormonism", "Scientology", "Unity", "None of the above"] },
      { id: "occultNewAge", type: "checkboxes", label: "Occult & New Age - Are you or have you been involved in any of the following? Select all that apply.", required: true,
        options: ["Astrology", "Crystals", "Fortune telling", "Magic 8 Ball", "Ouija Board", "Seances", "Tarot cards", "Hex", "Levitation games", "Blood sacrifice", "Sorcery", "Incantation", "Spells", "Witchcraft", "Shamanism", "Ancestor worship", "Yoga", "New Age/Eastern meditation", "None of the above"] },
      { id: "occultObjects", type: "radio", label: "Do you have any objects in your home or possession associated with other religions, cults, the occult, Wicca, Satanism?", required: true, options: YES_NO },
      { id: "occultObjectsExplain", type: "textarea", label: "Please explain", required: true,
        visibleIf: { id: "occultObjects", equals: "Yes" } },
      { id: "secretSocieties", type: "checkboxes", label: "Secret societies/associations - Have you been involved in any of the following? Select all that apply.", required: true,
        options: ["Eastern Star", "Elk Lodge", "Freemasonry", "Oddfellow", "Rainbow Girls", "Rebecca Lodge", "Shriners", "None of the above"] },
      { id: "ceremoniesCovenants", type: "radio", label: "Have you participated in any ceremonies or covenants associated with any of the above?", required: true, options: YES_NO },
      { id: "ceremoniesCovenantsExplain", type: "textarea", label: "Please explain", required: true,
        visibleIf: { id: "ceremoniesCovenants", equals: "Yes" } }
    ]
  },
  // ---------------------------------------------------------------- 10
  {
    id: "supernatural",
    title: "Supernatural experiences",
    questions: [
      { id: "supernaturalEncounter", type: "radio", label: "Have you ever had a supernatural encounter (seeing things or feeling an evil presence)?", required: true, options: YES_NO },
      { id: "supernaturalEncounterExplain", type: "textarea", label: "Please explain", required: true,
        visibleIf: { id: "supernaturalEncounter", equals: "Yes" } },
      { id: "supernaturalSymptoms", type: "checkboxes", label: "Do you experience any of the following? Select all that apply", required: true,
        options: ["Dreams of abuse or other sexual perversion", "Perverted thoughts or images during worship", "Sleep paralysis", "None of the above"] }
    ]
  },
  // ---------------------------------------------------------------- 11
  {
    id: "curses",
    title: "Curses",
    intro: "A black magic curse can be spoken or it can be an action. It calls on demonic power to inflict harm (breakdown of health, death, financial ruin, relational breakdown etc) on the recipient.",
    questions: [
      { id: "curseSpoken", type: "radio", label: "Has anyone spoken or preformed a black magic curse over you or over one of your family or ancestors?", required: true, options: YES_NO },
      { id: "curseSpokenExplain", type: "textarea", label: "Please explain", required: true,
        visibleIf: { id: "curseSpoken", equals: "Yes" } }
    ]
  },
  // ---------------------------------------------------------------- 12
  {
    id: "painfulExperiences",
    title: "Painful experiences",
    questions: [
      { id: "violence", type: "checkboxes", label: "Violence - Have you experienced any of the following either as a victim or a perpetrator? Select all that apply.", required: true,
        options: ["Accident", "Emotional abuse", "Jail/incarceration", "Kidnapping", "Massacre", "Physical abuse", "Physical assault", "Sexual abuse", "Surgery", "War", "None of the above"] },
      { id: "socialWounds", type: "checkboxes", label: "Social wounds - Have you experienced any of the following? Select all that apply.", required: true,
        options: ["Abandonment", "Career disappointment", "Family brokenness", "Neglect", "Rejection by family or friends or colleagues", "Unexpected loss of a loved one", "None of the above"] },
      { id: "negativeEmotions", type: "checkboxes", label: "Persistent negative emotions - Do you presently experience any of the following? Select all that apply.", required: true,
        options: ["Anger at God", "Anger at others", "Anger at self", "Anxiety/worry", "Bitterness", "Fear", "Fear - Losing your mind, death, accident, other", "Feelings of inadequacy", "Feelings of inferiority", "Feelings of unworthiness", "Grief", "Hatred", "Hopelessness", "Hurt from being abused", "Hurt from being rejected/neglected/abandoned", "Hurt from feeling unloved", "Hurt from church", "Rage", "Regret", "Resentment", "Sorrow", "Sadness", "None of the above"] },
      { id: "addictions", type: "checkboxes", label: "Addiction - Have you had or do you have any addiction(s)? Select all that apply.", required: true,
        options: ["Alcohol", "Food", "Smoking tobacco", "Gambling", "Porn", "Prescription drugs/pain meds", "Sex", "Street drugs", "Video games", "None of the above"] }
    ]
  },
  // ---------------------------------------------------------------- 13
  {
    id: "pronouncements",
    title: "Pronouncements and Inner Vows",
    intro: `Pronouncements. A pronouncement does not call on demonic powers and does not necessarily have the intention to inflict harm. It can be uttered either by a significant person in our life or ourself. Pronouncements made by a significant person, often a person in authority like a parent, teacher, employer/boss, can be uttered in anger or frustration or condemnation or disdain. Examples include statements such as: "You're stupid"; "You'll never amount to anything"; "You're fat"; "You always mess up"; "I don't know why we had you"; "You can't be in our group". Pronouncements we make can be spoken in discouragement, despair, frustration, anger, and even in joking self-disparagement. "I'm stupid…fat…lazy…dumb…ugly." You get the idea.`,
    questions: [
      { id: "pronouncementBySomeoneElse", type: "radio", label: "Has any significant person in your life (parent, teacher, boss etc.) ever spoken a pronouncement about you?", required: true, options: YES_NO },
      { id: "pronouncerRelationship", type: "textarea", label: "Who is that person (relationship to you)?", required: true,
        visibleIf: { id: "pronouncementBySomeoneElse", equals: "Yes" } },
      { id: "pronouncementText", type: "textarea", label: "What is/are the pronouncement(s)?", required: true,
        visibleIf: { id: "pronouncementBySomeoneElse", equals: "Yes" } },
      { id: "pronouncementBySelf", type: "radio", label: "Have you made a pronouncement about yourself?", required: true, options: YES_NO },
      { id: "pronouncementSelfText", type: "textarea", label: "What is/are the pronouncement(s)?", required: true,
        visibleIf: { id: "pronouncementBySelf", equals: "Yes" } }
    ]
  },
  // ---------------------------------------------------------------- 14
  {
    id: "innerVows",
    title: "Inner vows",
    intro: `An inner vow is strong decision, oath, or declaration of what we will or will not do to protect ourselves from pain or further hurt or to get what we think we need. Examples include statements like: "I'll never be like my mother"; "I'll never be poor"; "I'll never marry"; "I'll never cry"; "I'll always be in control".`,
    questions: [
      { id: "innerVowMade", type: "radio", label: "Have you ever made an inner vow?", required: true, options: YES_NO },
      { id: "innerVowText", type: "textarea", label: "What is/are the vow(s)?", required: true,
        visibleIf: { id: "innerVowMade", equals: "Yes" } }
    ]
  },
  // ---------------------------------------------------------------- 15
  {
    id: "personalSin",
    title: "Personal Sin & transgressions",
    intro: `The Bible uses different Hebrew words to describe different types of sin. "Transgression" refers to sin that we knowingly commit. "Sin" refers both to a general category of all sin as well as sin that we unknowingly commit. "Iniquity" refers to generational sin, that is sin committed not by us but by our ancestors, living and dead. The prayer team will help you deal with any sin that may have created an open door for the enemy to harass you.`,
    questions: [
      { id: "unconfessedSin", type: "radio", label: "Is there any unconfessed sin (transgression) in your life?", required: true, options: YES_NO },
      { id: "unconfessedSinExplain", type: "textarea", label: "Please explain", required: true,
        visibleIf: { id: "unconfessedSin", equals: "Yes" } }
    ]
  },
  // ---------------------------------------------------------------- 16
  {
    id: "unforgiveness",
    title: "Unforgiveness",
    intro: `One major block for healing is unforgiveness. Is there anyone who has hurt you that you have not forgiven – truly forgiven? Forgiving someone doesn't mean that you're saying that what they did was right or that it didn't have a negative impact on you. Forgiveness doesn't mean that the relationship with them will be restored – that's reconciliation and reconciliation may or may not be possible. Forgiveness is giving up the right/desire to expect someone pay for what they did. It's separate from legal consequences. Forgiveness is only about your attitude. It's not about a feeling. Rather, it's an act of your will. Even though forgiveness can be very difficult, the reality is that unforgiveness forms a prison for you. There's an easy way to test whether you need to forgive anyone. Simply think of the people who have hurt you. If you have a reaction of hurt, resentment, anger, bitterness or any other negative emotion when you think of a particular person, you haven't forgiven them.`,
    questions: [
      { id: "needToForgive", type: "radio", label: "Is there anyone you need to forgive?", required: true, options: ["Yes", "No", "Unsure"] },
      { id: "needToForgiveExplain", type: "textarea", label: "Please explain", required: true,
        visibleIf: { id: "needToForgive", in: ["Yes", "Unsure"] } },
      { id: "notForgivenSelf", type: "radio", label: "Is there something you've done or not done, said or not said, some failure, some sin, that you regret so much that you have not forgiven yourself?", required: true, options: YES_NO },
      { id: "notForgivenSelfExplain", type: "textarea", label: "Please explain", required: true,
        visibleIf: { id: "notForgivenSelf", equals: "Yes" } },
      { id: "angryAtGod", type: "radio", label: "Do you blame God, are you angry at God, for something He has done or not done?", required: true, options: YES_NO },
      { id: "angryAtGodExplain", type: "textarea", label: "Please explain.", required: true,
        visibleIf: { id: "angryAtGod", equals: "Yes" } }
    ]
  }
];

// Exposed as a single global for plain <script> includes (no build step,
// matches the rest of this project's single-file, no-framework
// convention) — everything above stays scoped inside this IIFE so it
// can't collide with same-named consts destructured from this object in
// the pages that consume it.
window.FPHM_INTAKE_CONFIG = { INTRO_TEXT, LIABILITY_TEXT, SECTIONS };

})();
